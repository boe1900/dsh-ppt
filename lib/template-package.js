import { createHash } from 'node:crypto';
import { lstat, readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { zipSync } from 'fflate';
import sharp from 'sharp';
import yaml from 'js-yaml';
import { SaxesParser } from 'saxes';
import { parse as parseCss, walk as walkCss, ident as cssIdent } from 'css-tree';
import { parsePptdProject, checkPptdProject } from './pptd.js';
import { validPersonalTemplateProfile, PERSONAL_TEMPLATE_PROFILE_VERSION } from './personal-template-profile.js';
import { validPersonalTemplateSamples } from './personal-template-samples.js';
import { parseTemplateDesign } from './template-design.js';

export const MAX_TEMPLATE_PACKAGE_BYTES = 60 * 1024 * 1024;
const MAX_EXPANDED = 256 * 1024 * 1024;
const MAX_PROJECT = 128 * 1024 * 1024;
const MAX_FILES = 1300;
const MAX_MANIFEST = 8 * 1024 * 1024;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fail = message => { throw new Error(`模板包无效：${message}`); };
const text = bytes => new TextDecoder('utf-8', { fatal: true }).decode(bytes);
const canonical = value => JSON.stringify(value, (_key, item) => item && typeof item === 'object' && !Array.isArray(item)
  ? Object.fromEntries(Object.keys(item).sort().map(key => [key, item[key]])) : item);

function safePath(value) {
  if (typeof value !== 'string' || value.length > 240 || /[\\:\u0000-\u001f\u007f]/u.test(value)
    || value.split('/').some(part => !part || part === '.' || part === '..' || /[. ]$/u.test(part))) fail('文件路径不安全');
  return value;
}
function fileLimit(file) {
  if (file === 'manifest.json') return MAX_MANIFEST;
  if (file === 'project/deck.pptd') return 512 * 1024;
  if (file === 'project/template-design.json') return 64 * 1024;
  if (/^project\/pages\/[A-Za-z0-9_.-]+\.page$/u.test(file)) return 2 * 1024 * 1024;
  if (/^project\/.+\.(png|jpe?g|gif|webp|svg)$/iu.test(file)) return 32 * 1024 * 1024;
  if (/^preview\/pages\/page-[1-9][0-9]{0,2}\.png$/u.test(file)) return 16 * 1024 * 1024;
  if (file === 'original/source.pptx') return MAX_PROJECT;
  fail(`不支持的文件 ${file}`);
}
const crcTable = Array.from({ length: 256 }, (_, value) => {
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});
function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function checkExtra(bytes) {
  for (let offset = 0; offset < bytes.length;) {
    if (offset + 4 > bytes.length) fail('ZIP 扩展字段损坏');
    const tag = bytes.readUInt16LE(offset), size = bytes.readUInt16LE(offset + 2);
    if (tag === 1 || offset + 4 + size > bytes.length) fail('不支持 ZIP64 或损坏的扩展字段');
    offset += 4 + size;
  }
}

function unpack(bytes) {
  if (bytes.length > MAX_TEMPLATE_PACKAGE_BYTES || bytes.length < 22) fail('文件大小超过 60 MiB 或 ZIP 不完整');
  let end = bytes.length - 22;
  while (end >= Math.max(0, bytes.length - 65557) && bytes.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0 || end < bytes.length - 65557 || bytes.readUInt32LE(end) !== 0x06054b50) fail('缺少 ZIP 目录');
  const count = bytes.readUInt16LE(end + 10), start = bytes.readUInt32LE(end + 16);
  if (bytes.readUInt16LE(end + 4) || bytes.readUInt16LE(end + 6) || bytes.readUInt16LE(end + 8) !== count
    || !count || count > MAX_FILES || start + bytes.readUInt32LE(end + 12) !== end
    || end + 22 + bytes.readUInt16LE(end + 20) !== bytes.length) fail('ZIP 目录或文件数量无效');
  const entries = [], names = new Set();
  let offset = start, expanded = 0;
  for (let index = 0; index < count; index++) {
    if (offset + 46 > end || bytes.readUInt32LE(offset) !== 0x02014b50) fail('ZIP 目录损坏');
    const flags = bytes.readUInt16LE(offset + 8), method = bytes.readUInt16LE(offset + 10);
    const crc = bytes.readUInt32LE(offset + 16), compressed = bytes.readUInt32LE(offset + 20), size = bytes.readUInt32LE(offset + 24);
    const length = bytes.readUInt16LE(offset + 28), extra = bytes.readUInt16LE(offset + 30), comment = bytes.readUInt16LE(offset + 32);
    const local = bytes.readUInt32LE(offset + 42), attrs = bytes.readUInt32LE(offset + 38);
    const type = attrs >>> 16 & 0xf000;
    if (offset + 46 + length + extra + comment > end || flags & ~0x800 || ![0, 8].includes(method)
      || bytes.readUInt16LE(offset + 34) || (type && type !== 0x8000) || attrs & 0x10) fail('不支持的 ZIP 文件类型、加密或压缩方式');
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + length), file = safePath(text(nameBytes));
    const key = file.normalize('NFC').toLowerCase();
    if (names.has(key)) fail('ZIP 包含重复或大小写冲突路径');
    names.add(key);
    if (size > fileLimit(file) || (expanded += size) > MAX_EXPANDED || compressed > MAX_TEMPLATE_PACKAGE_BYTES) fail('解压大小超过限制');
    checkExtra(bytes.subarray(offset + 46 + length, offset + 46 + length + extra));
    if (local + 30 > start || bytes.readUInt32LE(local) !== 0x04034b50
      || bytes.readUInt16LE(local + 6) !== flags || bytes.readUInt16LE(local + 8) !== method
      || bytes.readUInt32LE(local + 14) !== crc || bytes.readUInt32LE(local + 18) !== compressed
      || bytes.readUInt32LE(local + 22) !== size || bytes.readUInt16LE(local + 26) !== length) fail('ZIP 文件头不一致');
    const localExtra = bytes.readUInt16LE(local + 28), data = local + 30 + length + localExtra;
    if (data + compressed > start || !bytes.subarray(local + 30, local + 30 + length).equals(nameBytes)) fail('ZIP 文件范围无效');
    checkExtra(bytes.subarray(local + 30 + length, data));
    entries.push({ file, local, data, compressed, size, method, crc });
    offset += 46 + length + extra + comment;
  }
  if (offset !== end) fail('ZIP 目录长度不一致');
  let cursor = 0;
  for (const entry of [...entries].sort((a, b) => a.local - b.local)) {
    if (entry.local !== cursor) fail('ZIP 文件范围重叠或包含未声明数据');
    cursor = entry.data + entry.compressed;
  }
  if (cursor !== start) fail('ZIP 包含未声明数据');
  const files = new Map();
  for (const entry of entries) {
    const input = bytes.subarray(entry.data, entry.data + entry.compressed);
    const output = entry.method === 0 ? input : inflateRawSync(input, { maxOutputLength: Math.max(1, entry.size) });
    if (output.length !== entry.size || crc32(output) !== entry.crc) fail(`文件校验失败 ${entry.file}`);
    files.set(entry.file, output);
  }
  return files;
}

export function decodeTemplatePackage(input) {
  if (typeof input?.fileName !== 'string' || !/^[^/\\:\u0000-\u001f]{1,180}\.dshtpl$/iu.test(input.fileName)) fail('请选择 .dshtpl 模板包');
  const base64 = input?.base64;
  if (typeof base64 !== 'string' || base64.length > Math.ceil(MAX_TEMPLATE_PACKAGE_BYTES / 3) * 4
    || !base64.length || base64.length % 4 || /[^A-Za-z0-9+/=]/u.test(base64)) fail('Base64 数据无效或文件超过 60 MiB');
  const bytes = Buffer.from(base64, 'base64');
  if (bytes.toString('base64') !== base64) fail('Base64 数据无效');
  const files = unpack(bytes);
  if (!files.has('manifest.json')) fail('缺少 manifest.json');
  const manifest = JSON.parse(text(files.get('manifest.json')));
  files.delete('manifest.json');
  if (manifest?.format !== 'dsh-template' || manifest.version !== 1) fail('不支持的模板包版本');
  if (!Array.isArray(manifest.files) || manifest.files.length !== files.size) fail('文件清单不匹配');
  const declared = new Set();
  for (const entry of manifest.files) {
    const file = safePath(entry?.path), bytes = files.get(file);
    if (declared.has(file) || !bytes || entry.size !== bytes.length || entry.sha256 !== hash(bytes)) fail(`文件清单校验失败 ${file}`);
    declared.add(file);
  }
  return { metadata: manifest, files };
}

function boundedData(value, depth = 0, budget = { nodes: 0 }) {
  if (++budget.nodes > 500000 || depth > 20) fail('元数据过于复杂');
  if (typeof value === 'string' && (value.length > 4000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value))) fail('元数据文字无效');
  if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > 1e9)) fail('元数据数字无效');
  if (value && typeof value === 'object') for (const item of Object.values(value)) boundedData(item, depth + 1, budget);
}
function fields(value, names) {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => !names.includes(key))) fail('元数据包含未知字段');
}
function behaviorMetadata(metadata, project, pageIndex) {
  const sourceKind = metadata.sourceKind;
  if (!['native-pptd', 'converted-pptx'].includes(sourceKind)) fail('模板来源类型无效');
  const templateProfile = metadata.templateProfile, templateSamples = metadata.templateSamples;
  if (templateProfile !== undefined) {
    boundedData(templateProfile);
    if (templateProfile.version !== PERSONAL_TEMPLATE_PROFILE_VERSION || !validPersonalTemplateProfile(templateProfile, pageIndex)) fail('模板结构索引无效');
    fields(templateProfile, ['version', 'canvas', 'palette', 'typography', 'pages']);
    fields(templateProfile.canvas, ['width', 'height']);
    if (templateProfile.canvas.width !== project.width || templateProfile.canvas.height !== project.height) fail('模板结构画布与工程不一致');
    fields(templateProfile.palette, ['background', 'surface', 'text', 'muted', 'accent', 'secondary']);
    fields(templateProfile.typography, ['title', 'body']);
    for (const style of Object.values(templateProfile.typography)) fields(style, ['fontFace', 'fontSize', 'color', 'bold']);
    for (const [index, page] of templateProfile.pages.entries()) {
      fields(page, ['slideNumber', 'file', 'role', 'roleBasis', 'slots']);
      if (page.slots.length > 2000) fail('模板结构对象过多');
      const ids = new Set();
      for (const slot of page.slots) {
        fields(slot, ['elementId', 'kind', 'role', 'action', 'bounds', 'fontFace', 'fontSize', 'color', 'bold', 'wrap', 'fit', 'textCapacity']);
        if (ids.has(slot.elementId) || !project.pages[index].elements.some(element => element.elementId === slot.elementId && element.elementType === slot.kind)) fail('模板结构引用不存在或重复的对象');
        ids.add(slot.elementId);
      }
    }
  }
  if (templateSamples !== undefined) {
    boundedData(templateSamples);
    if (!validPersonalTemplateSamples(templateSamples, project.pages.length)) fail('模板示例索引无效');
    fields(templateSamples, ['version', 'simplifiedTextCount', 'layouts', 'representativePages']);
    for (const page of templateSamples.layouts) fields(page, ['slideNumber', 'role', 'content', 'representative']);
  }
  return { sourceKind, ...(templateProfile === undefined ? {} : { templateProfile }), ...(templateSamples === undefined ? {} : { templateSamples }) };
}
function sourceDocument(bytes, file, refs) {
  const value = yaml.load(text(bytes), { schema: yaml.JSON_SCHEMA });
  const seen = new Set();
  let nodes = 0;
  function inspect(item, depth = 0) {
    if (++nodes > 100000 || depth > 40) fail(`工程文本过于复杂 ${file}`);
    if (!item || typeof item !== 'object') return;
    if (seen.has(item)) fail(`工程不能包含 YAML 别名 ${file}`);
    seen.add(item);
    for (const [key, child] of Object.entries(item)) {
      if (key === 'src') refs.add(safePath(child));
      inspect(child, depth + 1);
    }
  }
  inspect(value);
  return value;
}
async function safeSvg(bytes) {
  const embedded = new Map();
  const reference = value => {
    if (value.startsWith('#')) return;
    const match = /^data:image\/(png|jpeg|gif|webp);base64,([A-Za-z0-9+/=]+)$/iu.exec(value);
    if (!match) fail('SVG 不能引用外部资源');
    const encoded = match[2], image = Buffer.from(encoded, 'base64');
    if (image.length > 32 * 1024 * 1024 || image.toString('base64') !== encoded) fail('SVG 内嵌图片无效');
    embedded.set(value, { image, extension: match[1].toLowerCase() === 'jpeg' ? 'jpg' : match[1].toLowerCase() });
  };
  const css = (value, context) => {
    const ast = parseCss(value, { context, parseCustomProperty: true, onParseError: error => { throw error; } });
    walkCss(ast, node => {
      if (node.type === 'Url') reference(node.value);
      if (node.type === 'Raw' || (node.type === 'Atrule' && cssIdent.decode(node.name).toLowerCase() === 'import')
        || (node.type === 'Function' && ['url', 'image', 'image-set', '-webkit-image-set', 'src'].includes(cssIdent.decode(node.name).toLowerCase()))) fail('SVG 样式不能引用外部资源');
    });
  };
  const presentation = new Set(['fill', 'stroke', 'filter', 'clip-path', 'mask', 'cursor', 'marker', 'marker-start', 'marker-mid', 'marker-end', 'color-profile']);
  const parser = new SaxesParser({ xmlns: true });
  const styles = [];
  parser.on('doctype', () => fail('SVG 不能包含外部实体'));
  parser.on('processinginstruction', instruction => { if (instruction.target.toLowerCase() === 'xml-stylesheet') fail('SVG 不能引用外部样式'); });
  parser.on('opentag', tag => {
    if (['script', 'foreignobject', 'iframe', 'object'].includes(tag.local.toLowerCase())) fail('SVG 包含可执行内容');
    if (tag.local.toLowerCase() === 'style') styles.push('');
    for (const attr of Object.values(tag.attributes)) {
      const name = attr.local.toLowerCase();
      if (/^on/u.test(name)) fail('SVG 包含可执行内容');
      if (name === 'href' || name === 'src') reference(attr.value);
      else if (name === 'style') css(attr.value, 'declarationList');
      else if (presentation.has(name)) css(attr.value, 'value');
    }
  });
  const collectStyle = value => { if (styles.length) styles[styles.length - 1] += value; };
  parser.on('text', collectStyle); parser.on('cdata', collectStyle);
  parser.on('closetag', tag => { if (tag.local.toLowerCase() === 'style') css(styles.pop(), 'stylesheet'); });
  parser.write(text(bytes)).close();
  for (const { image, extension } of embedded.values()) await validateImage(image, `embedded.${extension}`);
}
async function validateImage(bytes, file, project) {
  const preview = file.startsWith('preview/');
  if (/\.svg$/iu.test(file)) await safeSvg(bytes);
  const image = sharp(bytes, { limitInputPixels: preview ? 4000000 : 40000000, failOn: 'warning', animated: true });
  const info = await image.metadata();
  const expected = preview ? ['png'] : { '.png': ['png'], '.jpg': ['jpeg'], '.jpeg': ['jpeg'], '.gif': ['gif'], '.webp': ['webp'], '.svg': ['svg'] }[path.extname(file).toLowerCase()];
  if (!expected?.includes(info.format) || !info.width || !info.height || (info.pages ?? 1) > 200
    || (preview && (info.pages ?? 1) !== 1)
    || (preview && (Math.max(info.width, info.height) > 1920 || Math.abs(info.width / info.height - project.width / project.height) > 2 / info.height))) fail(`图片格式、尺寸或比例无效 ${file}`);
  await image.raw().toBuffer();
  return info.format === 'svg' ? 'image/svg+xml' : `image/${info.format}`;
}

export async function validateTemplatePackage({ metadata, files }, { previews = true, maxSlides = 200 } = {}) {
  if (!files.has('project/deck.pptd') || files.size > MAX_FILES - 1) fail('缺少模板工程或文件过多');
  let expanded = 0, projectBytes = 0;
  const aliases = new Set();
  for (const [file, bytes] of files) {
    safePath(file);
    const alias = file.normalize('NFC').toLowerCase();
    if (aliases.has(alias)) fail('文件路径大小写冲突');
    aliases.add(alias);
    if (bytes.length > fileLimit(file) || (expanded += bytes.length) > MAX_EXPANDED) fail('文件大小超过限制');
    if (file.startsWith('project/') && (projectBytes += bytes.length) > MAX_PROJECT) fail('工程超过 128 MiB');
  }
  const refs = new Set(), raw = sourceDocument(files.get('project/deck.pptd'), 'deck.pptd', refs);
  if (!Array.isArray(raw?.pages) || !raw.pages.length || raw.pages.length > maxSlides) fail('模板页面数量无效');
  const pages = new Map(), assets = new Map(), required = new Set(['project/deck.pptd']);
  for (const file of raw.pages) {
    safePath(file);
    if (!/^pages\/[A-Za-z0-9_.-]+\.page$/u.test(file) || pages.has(file)) fail('工程页面路径无效或重复');
    const bytes = files.get(`project/${file}`);
    if (!bytes) fail(`缺少页面 ${file}`);
    sourceDocument(bytes, file, refs);
    pages.set(file, text(bytes)); required.add(`project/${file}`);
  }
  for (const ref of refs) {
    const file = `project/${ref}`, bytes = files.get(file);
    if (!bytes || !/\.(png|jpe?g|gif|webp|svg)$/iu.test(ref)) fail(`缺少本地图片 ${ref}`);
    const mediaType = await validateImage(bytes, file);
    assets.set(ref, { path: ref, bytes, mediaType, sha256: hash(bytes) }); required.add(file);
  }
  const project = parsePptdProject({ entryName: 'deck.pptd', manifest: text(files.get('project/deck.pptd')), pages, assets });
  const check = checkPptdProject(project);
  if (check.status === 'fail') fail(`工程校验失败：${check.issues.filter(issue => issue.severity === 'error').slice(0, 5).map(issue => issue.message).join('；')}`);
  const pageIndex = project.pages.map((page, index) => ({ slideNumber: index + 1, file: page.file }));
  if (files.has('project/template-design.json')) {
    parseTemplateDesign(files.get('project/template-design.json'), project.pages); required.add('project/template-design.json');
  }
  if (files.has('original/source.pptx')) required.add('original/source.pptx');
  for (let page = 1; page <= project.pages.length; page++) {
    const file = `preview/pages/page-${page}.png`;
    if (previews) {
      if (!files.has(file)) fail(`缺少页面预览 ${page}`);
      await validateImage(files.get(file), file, project);
    }
    if (files.has(file)) required.add(file);
  }
  if (files.size !== required.size) fail('模板包包含未使用或未知文件');
  const behavior = behaviorMetadata(metadata, project, pageIndex);
  const digest = createHash('sha256').update('dsh-template-package-identity-v1\0').update(canonical(behavior)).update('\0');
  for (const [file, bytes] of [...files].filter(([file]) => !file.startsWith('preview/')).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0))
    digest.update(JSON.stringify([file, bytes.length])).update('\0').update(bytes);
  return { metadata: { name: metadata.name, description: metadata.description ?? '', ...behavior }, files, project, pageIndex, check, fingerprint: digest.digest('hex') };
}

export async function readTemplatePackage(directory, record, options) {
  const files = new Map();
  let total = 0;
  async function add(relative, file) {
    const info = await lstat(file);
    if (!info.isFile() || info.isSymbolicLink() || info.size > fileLimit(relative)
      || (total += info.size) > MAX_EXPANDED || files.size >= MAX_FILES - 1) fail('本地模板文件无效或超过大小限制');
    files.set(relative, await readFile(file));
  }
  async function walk(relative) {
    let folder = directory;
    for (const part of ['', ...relative.split('/')]) {
      folder = path.join(folder, part);
      const info = await lstat(folder);
      if (!info.isDirectory() || info.isSymbolicLink()) fail('本地模板目录无效');
    }
    for (const entry of await readdir(folder, { withFileTypes: true })) {
      const file = safePath(`${relative}/${entry.name}`);
      if (entry.isDirectory()) await walk(file);
      else await add(file, path.join(directory, file));
    }
  }
  await walk('project');
  if (options?.previews !== false) await walk('preview/pages');
  try { await add('original/source.pptx', path.join(directory, 'source.pptx')); }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  return validateTemplatePackage({ metadata: { name: record.name, description: record.description,
    sourceKind: record.authoringKind === 'native-pptd' ? 'native-pptd' : 'converted-pptx',
    templateProfile: record.templateProfile, templateSamples: record.templateSamples }, files }, options);
}

export function encodeTemplatePackage(snapshot) {
  const manifest = { format: 'dsh-template', version: 1, ...snapshot.metadata, files: [...snapshot.files].map(([file, bytes]) => ({ path: file, size: bytes.length, sha256: hash(bytes) })) };
  const manifestBytes = Buffer.from(JSON.stringify(manifest));
  if (manifestBytes.length > MAX_MANIFEST) fail('模板元数据超过大小限制');
  const entries = Object.fromEntries([['manifest.json', manifestBytes], ...snapshot.files]);
  const bytes = Buffer.from(zipSync(entries, { level: 6 }));
  if (bytes.length > MAX_TEMPLATE_PACKAGE_BYTES) fail('导出文件超过 60 MiB');
  return { fileName: `${snapshot.metadata.name.replace(/[<>:"/\\|?*\u0000-\u001f]/gu, '_').replace(/[. ]+$/u, '') || 'template'}.dshtpl`, base64: bytes.toString('base64') };
}
