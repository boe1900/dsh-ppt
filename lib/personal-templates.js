import { createHash, randomUUID } from 'node:crypto';
import { cp, lstat, mkdir, readFile, readdir, realpath, rename, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import sharp from 'sharp';
import { runPersonalTemplateImport } from './personal-template-import.js';
import { loadPptdProject, checkPptdProject, parsePptdProject } from './pptd.js';
import { renderOfficeTemplatePreview } from './office-template-preview.js';
import { prepareOfficeFontAccess } from './office-font-access.js';
import { extractPersonalTemplateProfile, personalTemplateSourceTheme, validPersonalTemplateProfile, PERSONAL_TEMPLATE_PROFILE_VERSION } from './personal-template-profile.js';
import { createPersonalTemplateSamples, validPersonalTemplateSamples } from './personal-template-samples.js';
import { PersonalTemplateUploads, personalTemplateFileName, decodeTemplateBase64 } from './personal-template-upload.js';
import { readTemplateDesign } from './template-design.js';
import { decodeTemplatePackage, validateTemplatePackage, readTemplatePackage, encodeTemplatePackage } from './template-package.js';

export const PERSONAL_TEMPLATE_CONVERSION_VERSION = 10;
export const PERSONAL_TEMPLATE_PREVIEW_VERSION = 11;
const MAX_TEMPLATES = 100;
// Template input pages have their own processing budget; generation remains capped separately.
export const MAX_PERSONAL_TEMPLATE_SLIDES = 200;
const PREVIEW_LONG_EDGE = 1920;
const THUMBNAIL_LONG_EDGE = 720;
// Bounds each JSON RPC request. Browser uploads transfer independent 8 MiB chunks.
export const MAX_PERSONAL_TEMPLATE_HTTP_BODY_BYTES = 86 * 1024 * 1024;
const ID = /^personal-[a-f0-9]{64}$/;
const DRAFT = /^[a-f0-9-]{36}$/;
const VERSION = /^(?:[a-f0-9]{64}|[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12})$/;
const COMMIT = /^[a-f0-9]{64}$/;
const conflict = () => new Error('原模板已更新，请重新打开或另存为新模板');
const baseOf = record => ({ templateId: record.id, version: record.version });
const hash = value => createHash('sha256').update(value).digest('hex');
const profileMetadata = templateProfile => ({templateProfile,palette:templateProfile.palette,
  titleFontFace:templateProfile.typography.title.fontFace,bodyFontFace:templateProfile.typography.body.fontFace,
  aspectRatio:templateProfile.canvas.width/templateProfile.canvas.height>=1.5?'wide':'standard'});
const nameOf = value => {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > 80 || /[\u0000-\u001f\u007f]/u.test(value))
    throw new Error('模板名称应为 1–80 个字符');
  return value.trim();
};
const descriptionOf = value => {
  if (typeof value !== 'string' || value.length > 1000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value))
    throw new Error('模板描述应为 1000 个字符以内的文字');
  return value.trim();
};
const exists = async target => {
  try { await lstat(target); return true; }
  catch (error) { if (error.code === 'ENOENT') return false; throw error; }
};

function conversionError(check, fileName, sha256) {
  const issues = check.issues.filter(issue => issue.severity === 'error');
  const details = issues.map(issue => [
    issue.page === undefined ? '文稿' : `第 ${issue.page} 页`,
    issue.elementId ? `对象 ${issue.elementId}` : '',
    issue.message
  ].filter(Boolean).join(' · '));
  const error = new Error([
    `模板导入遇到 ${check.errorCount} 项转换问题：`,
    ...details,
    '请保留原 PPT 文件，提供这些详情以排查转换兼容性。'
  ].join('\n'));
  error.conversion = { fileName, sha256, check };
  return error;
}

/** One library per host-owned Desktop profile; caller-supplied paths never select a library. */
export class PersonalTemplateLibrary {
  tail = Promise.resolve();
  catalogTail = Promise.resolve();
  imports = new Map();
  #unavailable = new Map();
  fontAccess;
  constructor(store, convert, writeSource, maxSlides = MAX_PERSONAL_TEMPLATE_SLIDES) {
    this.store = store;
    this.root = path.join(store.root, 'personal-templates');
    this.convert = convert;
    this.writeSource = writeSource;
    this.maxSlides = maxSlides;
    this.uploads = new PersonalTemplateUploads(this);
  }
  async locked(work) {
    const pending = this.tail.catch(() => {}).then(work);
    this.tail = pending;
    return pending;
  }
  catalogLocked(work) {
    const pending = this.catalogTail.catch(() => {}).then(work);
    this.catalogTail = pending;
    return pending;
  }
  async directory(...parts) {
    await mkdir(this.store.root, { recursive: true, mode: 0o700 });
    const target = path.join(this.root, ...parts);
    const relative = path.relative(this.store.root, target);
    let current = this.store.root;
    for (const part of relative.split(path.sep)) {
      current = path.join(current, part);
      await mkdir(current, { mode: 0o700 }).catch(error => { if (error.code !== 'EEXIST') throw error; });
      const info = await lstat(current);
      if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('模板目录应为本地独立目录');
    }
    return target;
  }
  async checkedFile(directory, file) {
    const target = path.join(directory, file);
    const info = await lstat(target);
    if (!info.isFile() || info.isSymbolicLink()) throw new Error('模板文件应为普通本地文件');
    const actual = await realpath(target);
    if (!actual.startsWith((await realpath(this.root)) + path.sep)) throw new Error('模板文件超出个人库范围');
    return target;
  }
  async readRecord(id) {
    if (!ID.test(id)) throw new Error('个人模板标识无效');
    return this.catalogLocked(() => this.#readDirectoryRecord(path.join(this.root, 'saved', id), id));
  }
  async #readDirectoryRecord(directory, id, validateProject = false) {
    const info = await lstat(directory);
    if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('模板目录无效');
    const bytes = await readFile(await this.checkedFile(directory, 'template.json'));
    let record = JSON.parse(bytes.toString('utf8'));
    if (record.version !== undefined && !VERSION.test(record.version)) throw new Error('模板版本无效');
    record.version ??= hash(bytes);
    const payload = await this.#payloadDirectory(directory, record);
    if (!ID.test(record.id) || (id !== undefined && record.id !== id) || record.origin !== 'personal') throw new Error('模板记录无效');
    if (!Number.isInteger(record.slideCount) || record.slideCount < 1 || record.slideCount > this.maxSlides
      || typeof record.createdAt !== 'string' || !Number.isFinite(Date.parse(record.createdAt))) throw new Error('模板记录无效');
    if (!Array.isArray(record.previewImages) || record.previewImages.length !== 3 || record.previewImages.some(image =>
      typeof image !== 'string' || !image.startsWith('data:image/png;base64,') || image.length > 4 * 1024 * 1024)) throw new Error('模板预览资源无效');
    nameOf(record.name);
    if (!Array.isArray(record.pageIndex) || record.pageIndex.length !== record.slideCount || record.pageIndex.some((page, i) =>
      page.slideNumber !== i + 1 || typeof page.file !== 'string' || !/^pages\/[A-Za-z0-9_.-]+\.page$/.test(page.file))) throw new Error('模板页面索引无效');
    if(record.templateSamples!==undefined && !validPersonalTemplateSamples(record.templateSamples,record.slideCount)) throw new Error('模板示例索引无效，请检查模板工程');
    const invalidProfile = record.templateProfile !== undefined && !validPersonalTemplateProfile(record.templateProfile, record.pageIndex);
    if (invalidProfile && record.authoringKind !== 'native-pptd') throw new Error('模板结构索引无效，请重新导入模板');
    if (validateProject || invalidProfile) {
      const project = await this.#loadOwnedProject(path.join(payload, 'project'));
      if (project.pages.length !== record.pageIndex.length || project.pages.some((page, index) => page.file !== record.pageIndex[index].file))
        throw new Error('模板工程与页面索引不一致');
      if (validateProject) await readTemplateDesign(path.join(payload, 'project'), project.pages);
      if (invalidProfile) {
        const profile = extractPersonalTemplateProfile(project);
        if (!validPersonalTemplateProfile(profile, record.pageIndex)) throw new Error('模板结构索引生成失败，请检查模板工程');
        record = { ...record, ...profileMetadata(profile) };
      }
    }
    return record;
  }
  async #payloadDirectory(directory, record) {
    if (record.contentRevision === undefined) return directory;
    if (!DRAFT.test(record.contentRevision)) throw new Error('模板内容版本无效');
    let current = directory;
    for (const part of ['revisions', record.contentRevision]) {
      current = path.join(current, part);
      const info = await lstat(current);
      if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('模板内容目录无效');
    }
    return current;
  }
  async #recordPayload(record) {
    return this.#payloadDirectory(path.join(this.root, 'saved', record.id), record);
  }
  async #availableId(preferred) {
    const saved = await this.directory('saved');
    let id = preferred;
    while (!id || await exists(path.join(saved, id))) id = `personal-${hash(randomUUID())}`;
    return id;
  }
  async #receipt(commitId) {
    if (!COMMIT.test(commitId)) throw new Error('模板保存回执无效');
    let receipt;
    try { receipt = JSON.parse(await readFile(await this.checkedFile(path.join(this.root, 'commit-receipts'), `${commitId}.json`), 'utf8')); }
    catch (error) { if (error.code === 'ENOENT') return undefined; throw error; }
    if (!['pending', 'committed'].includes(receipt.status) || receipt.commitId !== commitId || !COMMIT.test(receipt.requestDigest)
      || !ID.test(receipt.template?.id) || receipt.base?.templateId !== receipt.template.id || receipt.base.version !== receipt.template.version
      || !VERSION.test(receipt.base.version)) throw new Error('模板保存回执无效');
    return receipt;
  }
  async #writeReceipt(receipt) {
    await this.store.writeJson(path.join(await this.directory('commit-receipts'), `${receipt.commitId}.json`), receipt);
  }
  async #reconcileRecord(record) {
    if (record.lastCommitId === undefined) return;
    const receipt = await this.#receipt(record.lastCommitId);
    if (!receipt) throw new Error('模板保存回执缺失，请检查模板库');
    if (receipt.status === 'committed') return;
    if (receipt.template.id !== record.id || receipt.template.version !== record.version) throw new Error('模板保存回执与当前版本不一致');
    await this.#writeReceipt({ ...receipt, status: 'committed' });
  }
  list() {
    return this.catalogLocked(() => this.#readCatalog(true));
  }
  async #readCatalog(reportUnavailable) {
    const directory = path.join(this.root, 'saved');
    if (!await exists(directory)) { this.#unavailable.clear(); return []; }
    if ((await lstat(directory)).isSymbolicLink()) throw new Error('个人模板库目录无效');
    const items = await readdir(directory, { withFileTypes: true });
    const records = [];
    const unavailable = new Map();
    for (const item of items) {
      if (!ID.test(item.name)) continue;
      try { records.push(await this.#readDirectoryRecord(path.join(directory, item.name), item.name)); }
      catch (error) {
        const message = error.message;
        unavailable.set(item.name, message);
        if (reportUnavailable && this.#unavailable.get(item.name) !== message) await this.store.recordAudit('', {
          id: randomUUID(), operation: 'skip-unavailable-template', actor: 'user', status: 'failed',
          startedAt: new Date().toISOString(), completedAt: new Date().toISOString(), error: message
        }, { templateId: item.name });
      }
    }
    if (reportUnavailable) this.#unavailable = unavailable;
    return records.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }
  async audit(sessionId, operation, work) {
    const startedAt = new Date().toISOString();
    try {
      const result = await work();
      await this.store.recordAudit(sessionId, {
        id: randomUUID(), operation, actor: 'user', status: 'completed', startedAt,
        completedAt: new Date().toISOString(), summary: operation
      }, { templateId: result?.id, draftId: result?.draftId });
      return result;
    } catch (error) {
      await this.store.recordAudit(sessionId, {
        id: randomUUID(), operation, actor: 'user', status: 'failed', startedAt,
        completedAt: new Date().toISOString(), error: error.message
      }, error.conversion ? { conversion: error.conversion } : {});
      throw error;
    }
  }
  async uploadStart(sessionId, input) {
    return this.locked(() => this.audit(sessionId, 'upload-personal-template', async () => {
      const upload=await this.uploads.start(sessionId,input);
      try {await this.ensureFontAccess();return upload;}
      catch(error) {await this.uploads.cancel(sessionId,upload.uploadId);throw error;}
    }));
  }
  async ensureFontAccess() {
    // One check per host library. Optional denial retains the SDK's font fallbacks.
    const pending=this.fontAccess ??= prepareOfficeFontAccess();
    try {return await pending;}
    catch(error) {if(this.fontAccess===pending)this.fontAccess=undefined;throw error;}
  }
  async uploadChunk(sessionId, input) {
    return this.locked(() => this.audit(sessionId, 'upload-personal-template-chunk', () => this.uploads.append(sessionId, input)));
  }
  async cancelUpload(sessionId, uploadId) {
    if (!DRAFT.test(uploadId)) throw new Error('上传标识无效');
    const jobs = [...this.imports.values()].filter(job => job.uploadId === uploadId);
    const owned = jobs.filter(job => job.session === hash(sessionId));
    if (jobs.length && !owned.length) throw new Error('该处理任务属于其他会话');
    // Interrupt owned workers before joining the library queue that they hold.
    for (const job of owned) job.controller.abort();
    return this.locked(() => this.audit(sessionId, 'cancel-personal-template-upload', async () => {
      await this.uploads.cancel(sessionId, uploadId);
      for (const job of owned) {
        let directory;
        try { directory = await this.draft(sessionId, job.draftId); }
        catch (error) { if (error.code !== 'ENOENT') throw error; }
        if (directory) await rm(directory, { recursive: true, force: true });
      }
      return { uploadId };
    }));
  }
  async prepare(sessionId, input) {
    const requestId=input?.requestId ?? randomUUID();
    if(!DRAFT.test(requestId)) throw new Error('处理标识无效');
    for(const [id,job] of this.imports) if(job.status!=='processing' && Date.now()-job.startedAt>3600000) this.imports.delete(id);
    if(this.imports.has(requestId)||this.imports.size>=32) throw new Error('请稍后重新处理模板');
    const progress={session:hash(sessionId),startedAt:Date.now(),total:0,pages:[],status:'processing',
      uploadId:input?.uploadId,draftId:randomUUID(),controller:new AbortController()};
    this.imports.set(requestId,progress);
    try { return await this.locked(() => this.audit(sessionId, 'prepare-personal-template', async () => {
      progress.controller.signal.throwIfAborted();
      let uploadId=input?.uploadId;
      if(uploadId===undefined) {
        const fileName=personalTemplateFileName(input?.fileName),bytes=decodeTemplateBase64(input?.base64);
        const started=await this.uploads.start(sessionId,{fileName,size:bytes.length});uploadId=started.uploadId;
        try {
          for(let offset=0;offset<bytes.length;offset+=started.chunkBytes)
            await this.uploads.append(sessionId,{uploadId,offset,base64:bytes.subarray(offset,offset+started.chunkBytes).toString('base64')});
        } catch(error) {await this.uploads.cancel(sessionId,uploadId);throw error;}
      }
      const upload=await this.uploads.record(sessionId,uploadId);
      progress.uploadId=uploadId;
      const draftId=progress.draftId;
      if((await lstat(upload.file)).size!==upload.record.size)throw new Error('请完成文件上传后再处理模板');
      try {
        await this.ensureFontAccess();
        progress.controller.signal.throwIfAborted();
        const result=await runPersonalTemplateImport({root:this.store.root,sessionId,uploadId,draftId,maxSlides:this.maxSlides}, update=>{
          progress.total=update.total;
          if(update.page)progress.pages.push(update.page);
        }, {signal:progress.controller.signal});
        progress.controller.signal.throwIfAborted();
        if (result.duplicate && result.template.templateProfile) {
          const current = await this.readRecord(result.template.id);
          if (current.version !== result.template.version) throw conflict();
          if (current.templateProfile?.version !== PERSONAL_TEMPLATE_PROFILE_VERSION) {
            await this.#reconcileRecord(current);
            const template = { ...current, ...profileMetadata(result.template.templateProfile), version: randomUUID() };
            await this.catalogLocked(() => this.store.writeJson(path.join(this.root, 'saved', current.id, 'template.json'), template));
            return { duplicate: true, template };
          }
        }
        return result;
      } catch(error) {await rm(path.join(this.root,'drafts',draftId),{recursive:true,force:true});throw error;} finally {await rm(upload.directory,{recursive:true,force:true});}
    })); } finally { progress.status=progress.controller.signal.aborted?'cancelled':'finished';progress.startedAt=Date.now(); }
  }
  /** Runs inside the host-owned import process, under the parent's policy/audit lock. */
  async prepareSource(sessionId,input,progress,signal) {
      signal?.throwIfAborted();
      const started=Date.now();
      let upload;
      try {
        if (input?.uploadId !== undefined) upload = await this.uploads.read(sessionId, input.uploadId);
        const fileName = upload?.fileName ?? personalTemplateFileName(input?.fileName);
        const bytes = upload?.bytes ?? decodeTemplateBase64(input?.base64);
        if (bytes.length < 4 || !bytes.subarray(0, 4).equals(Buffer.from([80, 75, 3, 4])))
          throw new Error('请选择有效的 PPTX 文件');
        const sourceHash = hash(bytes), id = `personal-${sourceHash}`;
        const templates=await this.#readCatalog(false);
        let saved;
        for (const candidate of templates) {
          if (candidate.authoringKind === 'native-pptd' || candidate.sha256 !== sourceHash) continue;
          try {
            const payload = await this.#recordPayload(candidate);
            const original = await readFile(await this.checkedFile(payload, 'source.pptx'));
            if (hash(original) === sourceHash) { saved = candidate; break; }
          } catch { /* An unavailable historical source cannot prove this upload is a duplicate. */ }
        }
        if (saved?.conversionVersion === PERSONAL_TEMPLATE_CONVERSION_VERSION && saved?.previewVersion === PERSONAL_TEMPLATE_PREVIEW_VERSION) {
          if(saved.templateProfile?.version===PERSONAL_TEMPLATE_PROFILE_VERSION)return {duplicate:true,template:saved};
          // Enrich a current cached conversion without repeating Office rendering.
          const project=await loadPptdProject(await this.projectSource(saved.id));
          const profile=extractPersonalTemplateProfile(project);
          if(!validPersonalTemplateProfile(profile,saved.pageIndex)) throw new Error('模板结构索引生成失败，请检查模板工程');
          return {duplicate:true,template:{...saved,...profileMetadata(profile)}};
        }
        if (!saved && templates.length >= MAX_TEMPLATES) throw new Error('个人模板库最多保存 100 套模板');
        const drafts = await this.directory('drafts');
        // Expired previews are temporary resources, independent of registered templates.
        for (const item of await readdir(drafts, { withFileTypes: true })) {
          if (item.isDirectory() && DRAFT.test(item.name) && Date.now() - (await lstat(path.join(drafts, item.name))).mtimeMs > 86400000)
            await rm(path.join(drafts, item.name), { recursive: true });
        }
        if ((await readdir(drafts)).length >= 10) throw new Error('待保存模板达到 10 套，请先保存或取消已有预览');
        const conversionStarted=Date.now();
        const converted = await this.convert(bytes, fileName, {signal});
        signal?.throwIfAborted();
        const convertedAt=Date.now();
        if (!converted.slideCount || converted.slideCount > this.maxSlides) throw new Error(`模板应包含 1–${this.maxSlides} 页`);
        const missing = converted.diagnostics.filter(item => item.level === 'unsupported');
        if (missing.length) throw conversionError({ errorCount: missing.length, issues: missing.map(item => ({
          severity: 'error', code: 'conversion-unsupported', page: item.slide, elementId: item.nodeId,
          message: `${item.feature}：${item.message}`
        })) }, fileName, hash(bytes));
        progress?.({total:converted.slideCount});
        const original = parsePptdProject(converted.source);
        const checked = checkPptdProject(original);
        if (checked.status === 'fail') throw conversionError(checked, fileName, hash(bytes));
        const templateProfile=extractPersonalTemplateProfile(original);
        if(!validPersonalTemplateProfile(templateProfile,original.pages)) throw new Error('模板结构索引生成失败，请检查模板工程');
        const {source:sampleSource,samples:templateSamples}=createPersonalTemplateSamples(original,templateProfile);
        const source=personalTemplateSourceTheme(sampleSource,templateProfile);
        const parsed = parsePptdProject(source);
        const sampleCheck=checkPptdProject(parsed);
        if(sampleCheck.status==='fail')throw conversionError(sampleCheck,fileName,hash(bytes));
        const draftId = input.draftId ?? randomUUID();
        const directory = await this.directory('drafts', draftId);
        try {
          const project = await this.directory('drafts', draftId, 'project');
          await this.writeSource(project, source);
          await writeFile(path.join(directory, 'source.pptx'), bytes, { mode: 0o600 });
          const previewDirectory = path.join(directory, 'preview');
          signal?.throwIfAborted();
          const previewMetadata = await renderOfficeTemplatePreview(parsed,directory,PREVIEW_LONG_EDGE,this.maxSlides,progress ? async ({page,path:imagePath}) => {
            const thumbnail=await sharp(imagePath).resize({width:240,height:240,fit:'inside',withoutEnlargement:true}).png().toBuffer();
            progress({total:converted.slideCount,page:{page,preview:`data:image/png;base64,${thumbnail.toString('base64')}`}});
          } : undefined, signal);
          signal?.throwIfAborted();
          const previewPage = async page => readFile(path.join(previewDirectory, 'pages', `page-${page}.png`));
          const uniquePages=templateSamples.representativePages;
          const representativePages=[uniquePages[0],uniquePages[Math.floor(uniquePages.length/2)],uniquePages.at(-1)];
          const thumbnails=new Map();
          for(const page of new Set(representativePages)) thumbnails.set(page,(async()=>{
            const thumbnail = await sharp(path.join(previewDirectory,'pages',`page-${page}.png`)).resize({
              width: THUMBNAIL_LONG_EDGE, height: THUMBNAIL_LONG_EDGE, fit: 'inside', withoutEnlargement: true
            }).png().toBuffer();
            return `data:image/png;base64,${thumbnail.toString('base64')}`;
          })());
          const previewImages=await Promise.all(representativePages.map(page=>thumbnails.get(page)));
          const preview = `data:image/png;base64,${(await previewPage(1)).toString('base64')}`;
          const name = nameOf(fileName.replace(/\.pptx$/i, '').slice(0, 80));
          const record = {
            id, conversionVersion: PERSONAL_TEMPLATE_CONVERSION_VERSION, previewVersion: PERSONAL_TEMPLATE_PREVIEW_VERSION, origin: 'personal', name: saved?.name ?? name, category: 'personal', supportedModes: ['ppt'], aspectRatio: 'wide',
            description: `${converted.slideCount} 页个人 PPT 模板`, previewTitle: name, previewSubtitle: '我的模板',
            ...profileMetadata(templateProfile),
            templateSamples,
            createdAt: saved?.createdAt ?? new Date().toISOString(), ...(saved ? {description:saved.description,updatedAt:new Date().toISOString()} : {}), fileName, sha256: hash(bytes), slideCount: converted.slideCount,
            previewImages, previewMetadata,
            preparationTimings:{conversionMs:convertedAt-conversionStarted,totalMs:Date.now()-started},
            diagnostics: converted.diagnostics, checkWarnings: checked.warningCount,
            resourcePlaceholders: converted.diagnostics.filter(item => item.level === 'placeholder'),
            reviewIssues: checked.issues.filter(issue => issue.severity === 'warning').map(issue => ({
              ...issue, message: issue.code === 'out-of-bounds' ? '源文件对象位于页面边缘或画布外，预览按页面范围显示。'
                : issue.code === 'text-overflow' ? '源文件文字可能超出文本框，请在预览中核对。' : issue.message
            })),
            pageIndex: [...converted.source.pages.keys()].map((file, i) => ({ slideNumber: i + 1, file }))
          };
          await writeFile(path.join(directory, 'template.json'), JSON.stringify(record), { mode: 0o600 });
          await writeFile(path.join(directory, 'owner.json'), JSON.stringify({ session: hash(sessionId) }), { mode: 0o600 });
          signal?.throwIfAborted();
          return { draftId, template: record, preview };
        } catch (error) { await rm(directory, { recursive: true, force: true }); throw error; }
      } finally { if (upload) await rm(upload.directory, { recursive: true, force: true }); }
  }

  async importProgress(sessionId,requestId,after=0) {
    if(!DRAFT.test(requestId)||!Number.isInteger(after)||after<0||after>this.maxSlides) throw new Error('处理进度参数无效');
    const job=this.imports.get(requestId);
    if(!job) return {total:0,completed:0,pages:[],status:'waiting'};
    if(job.session!==hash(sessionId)) throw new Error('该处理任务属于其他会话');
    return {total:job.total,completed:job.pages.length,pages:job.pages.slice(after),status:job.status};
  }
  async draft(sessionId, draftId) {
    if (!DRAFT.test(draftId)) throw new Error('预览标识无效');
    const directory = path.join(this.root, 'drafts', draftId);
    const info = await lstat(directory);
    if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('预览目录无效');
    const owner = JSON.parse(await readFile(await this.checkedFile(directory, 'owner.json'), 'utf8'));
    if (owner.session !== hash(sessionId)) throw new Error('该预览属于其他会话');
    return directory;
  }
  async readPreview(directory,record,page) {
    if(!Number.isInteger(page)||page<1||page>record.slideCount)throw new Error('预览页码无效');
    const image=await this.checkedFile(path.join(directory,'preview','pages'),`page-${page}.png`);
    return { page, preview: `data:image/png;base64,${(await readFile(image)).toString('base64')}` };
  }
  async previewPage(sessionId,draftId,page) {
    return this.locked(()=>this.audit(sessionId,'preview-personal-template',async()=>{
      const directory=await this.draft(sessionId,draftId);
      const record=JSON.parse(await readFile(await this.checkedFile(directory,'template.json'),'utf8'));
      return this.readPreview(directory,record,page);
    }));
  }
  async previewSavedPage(sessionId, templateId, page) {
    return this.locked(() => this.audit(sessionId, 'preview-personal-template', async () => {
      const record = await this.readRecord(templateId);
      return this.readPreview(await this.#recordPayload(record),record,page);
    }));
  }
  async exportPackage(sessionId, templateId) {
    return this.locked(async () => {
      const record = await this.readRecord(templateId);
      return encodeTemplatePackage(await readTemplatePackage(await this.#recordPayload(record), record, { maxSlides: this.maxSlides }));
    });
  }
  async importPackage(sessionId, input, signal) {
    return this.locked(async () => {
      signal?.throwIfAborted();
      const snapshot = await validateTemplatePackage(decodeTemplatePackage(input), { maxSlides: this.maxSlides });
      const name = nameOf(snapshot.metadata.name), description = descriptionOf(snapshot.metadata.description);
      let id = `personal-${snapshot.fingerprint}`;
      const templates = await this.catalogLocked(() => this.#readCatalog(false));
      for (const record of templates) {
        try {
          const existing = await readTemplatePackage(await this.#recordPayload(record), record, { previews: false, maxSlides: this.maxSlides });
          if (existing.fingerprint === snapshot.fingerprint) return { outcome: 'existing', template: record };
        } catch (error) {
          if (record.id === id) throw new Error('模板库完整性冲突：目标模板无法验证，请检查本地模板');
        }
      }
      if (templates.some(record => record.id === id)) id = await this.#availableId();
      const destination = path.join(await this.directory('saved'), id);
      if (await exists(destination)) throw new Error('模板库完整性冲突：目标目录已存在，无法安全导入');
      if (templates.length >= MAX_TEMPLATES) throw new Error('个人模板库最多保存 100 套模板');
      const directory = await this.directory('package-staging', randomUUID());
      let published = false;
      try {
        for (const [file, bytes] of snapshot.files) {
          signal?.throwIfAborted();
          const relative = file === 'original/source.pptx' ? 'source.pptx' : file;
          const target = path.join(directory, relative);
          await mkdir(path.dirname(target), { recursive: true, mode: 0o700 });
          await writeFile(target, bytes, { mode: 0o600, flag: 'wx' });
        }
        const samples = snapshot.metadata.templateSamples;
        const representative = samples ? [samples.representativePages[0], samples.representativePages[Math.floor(samples.representativePages.length / 2)], samples.representativePages.at(-1)]
          : [1, Math.ceil(snapshot.pageIndex.length / 2), snapshot.pageIndex.length];
        const previewImages = [];
        for (const page of representative) {
          const bytes = await sharp(snapshot.files.get(`preview/pages/page-${page}.png`)).resize({ width: THUMBNAIL_LONG_EDGE, height: THUMBNAIL_LONG_EDGE, fit: 'inside', withoutEnlargement: true }).png().toBuffer();
          previewImages.push(`data:image/png;base64,${bytes.toString('base64')}`);
        }
        const profile = snapshot.metadata.templateProfile;
        const display = profileMetadata(profile ?? extractPersonalTemplateProfile(snapshot.project));
        delete display.templateProfile;
        const record = {
          id, version: randomUUID(), sha256: snapshot.fingerprint, conversionVersion: PERSONAL_TEMPLATE_CONVERSION_VERSION, previewVersion: PERSONAL_TEMPLATE_PREVIEW_VERSION,
          origin: 'personal', category: 'personal', supportedModes: ['ppt'], name, description,
          ...(snapshot.metadata.sourceKind === 'native-pptd' ? { authoringKind: 'native-pptd' } : {}),
          previewTitle: name, previewSubtitle: '我的模板', createdAt: new Date().toISOString(),
          fileName: snapshot.metadata.sourceKind === 'native-pptd' ? 'deck.pptd' : 'source.pptx',
          slideCount: snapshot.pageIndex.length, pageIndex: snapshot.pageIndex,
          aspectRatio: snapshot.project.width / snapshot.project.height >= 1.5 ? 'wide' : 'standard',
          ...display, ...(profile ? { templateProfile: profile } : {}), ...(samples ? { templateSamples: samples } : {}),
          previewImages, previewMetadata: { source: 'template-package', version: 1 },
          diagnostics: [], resourcePlaceholders: [], checkWarnings: snapshot.check.warningCount,
          reviewIssues: snapshot.check.issues.filter(issue => issue.severity === 'warning'),
        };
        await writeFile(path.join(directory, 'template.json'), JSON.stringify(record), { mode: 0o600, flag: 'wx' });
        await this.#readDirectoryRecord(directory, id, true);
        await this.catalogLocked(async () => {
          signal?.throwIfAborted();
          if (await exists(destination)) throw new Error('模板库完整性冲突：目标目录已存在，无法安全导入');
          await rename(directory, destination);
          published = true;
        });
        await this.store.recordAudit(sessionId, { id: randomUUID(), operation: 'import-template-package', actor: 'user', status: 'completed',
          startedAt: record.createdAt, completedAt: new Date().toISOString(), summary: 'import-template-package' }, { templateId: id }).catch(() => {});
        return { outcome: 'imported', template: record };
      } finally { if (!published) await rm(directory, { recursive: true, force: true }); }
    });
  }
  save(sessionId, draftId, name) {
    return this.locked(() => this.audit(sessionId, 'save-personal-template', async () =>
      (await this.#commitDraft(sessionId, draftId, { name: nameOf(name), destination: { kind: 'import' } })).template));
  }
  commit(sessionId, draftId, input) {
    const name = nameOf(input.name), destination = input.destination;
    if (!['copy', 'original'].includes(destination?.kind)
      || (destination.kind === 'original' && (!ID.test(destination.base?.templateId) || !VERSION.test(destination.base?.version))))
      throw new Error('模板保存目标无效');
    const target = destination.kind === 'copy' ? { kind: 'copy' } : { kind: 'original', base: { templateId: destination.base.templateId, version: destination.base.version } };
    return this.locked(() => this.audit(sessionId, 'commit-authored-template', () => this.#commitDraft(sessionId, draftId, { name, destination: target })));
  }
  async #commitDraft(sessionId, draftId, input) {
    if (!DRAFT.test(draftId)) throw new Error('模板保存标识无效');
    const commitId = hash(`${hash(sessionId)}:${draftId}`), requestDigest = hash(JSON.stringify(input));
    let receipt = await this.#receipt(commitId);
    if (receipt?.status === 'pending') {
      let record;
      try { record = await this.readRecord(receipt.template.id); }
      catch (error) { if (error.code !== 'ENOENT') throw error; }
      if (record?.lastCommitId === commitId) {
        await this.#reconcileRecord(record);
        receipt = await this.#receipt(commitId);
      }
    }
    if (receipt && receipt.requestDigest !== requestDigest) {
      if (receipt.status === 'committed') throw new Error('该预览已使用其他名称或保存目标提交，请生成新的预览');
      receipt = undefined;
    }
    if (receipt?.status === 'committed') return { template: receipt.template, base: receipt.base };
    const directory = await this.draft(sessionId, draftId);
    const draft = await this.#readDirectoryRecord(directory, undefined, true);
    const templates = await this.list();
    let original;
    if (input.destination.kind === 'original') {
      try { original = await this.readRecord(input.destination.base.templateId); }
      catch (error) { if (error.code === 'ENOENT') throw new Error('原模板已删除，可以另存为新模板'); throw error; }
      await this.#reconcileRecord(original);
      if (original.version !== input.destination.base.version) throw conflict();
    } else if (input.destination.kind === 'import') {
      const incoming = await readTemplatePackage(directory, draft, { previews: false, maxSlides: this.maxSlides });
      for (const record of templates) {
        let snapshot;
        try { snapshot = await readTemplatePackage(await this.#recordPayload(record), record, { previews: false, maxSlides: this.maxSlides }); }
        catch { continue; }
        if (snapshot.fingerprint !== incoming.fingerprint) continue;
        await this.#reconcileRecord(record);
        await this.#writeReceipt({ status: 'committed', commitId, requestDigest, template: record, base: baseOf(record) });
        await rm(directory, { recursive: true, force: true });
        return { template: record, base: baseOf(record) };
      }
    }
    if (!original && templates.length >= MAX_TEMPLATES) throw new Error('个人模板库最多保存 100 套模板');
    if (!receipt) {
      const id = original?.id ?? await this.#availableId(input.destination.kind === 'import' ? draft.id : undefined);
      const next = { ...draft, id, name: input.name, previewTitle: input.name,
        createdAt: original?.createdAt ?? new Date().toISOString(), updatedAt: new Date().toISOString(),
        version: randomUUID(), lastCommitId: commitId, ...(original ? { contentRevision: randomUUID() } : {}) };
      receipt = { status: 'pending', commitId, requestDigest, template: next, base: baseOf(next) };
      await this.#writeReceipt(receipt);
    }
    const next = receipt.template, destination = path.join(await this.directory('saved'), next.id);
    const staging = await this.directory('commit-staging', randomUUID());
    const stage = path.join(staging, 'payload');
    try {
      await cp(directory, stage, { recursive: true, errorOnExist: true, force: false });
      const payloadRecord = { ...next }; delete payloadRecord.contentRevision;
      await this.store.writeJson(path.join(stage, 'template.json'), payloadRecord);
      await this.#readDirectoryRecord(stage, next.id, true);
      if (original) {
        const revisions = await this.directory('saved', next.id, 'revisions');
        const payload = path.join(revisions, next.contentRevision);
        if (await exists(payload)) {
          const installed = await this.#readDirectoryRecord(payload, next.id, true);
          if (installed.version !== next.version || installed.lastCommitId !== commitId) throw new Error('模板内容版本冲突');
        } else await rename(stage, payload);
        await this.catalogLocked(() => this.store.writeJson(path.join(destination, 'template.json'), next));
      } else {
        await this.catalogLocked(async () => {
          if (await exists(destination)) throw new Error('模板保存目标已存在，请检查模板库');
          await rename(stage, destination);
        });
      }
      await this.#writeReceipt({ ...receipt, status: 'committed' });
      await rm(directory, { recursive: true, force: true }).catch(() => {});
      return { template: next, base: receipt.base };
    } finally { await rm(staging, { recursive: true, force: true }).catch(() => {}); }
  }
  async savedDraft(sessionId, draftId) {
    if (!DRAFT.test(draftId)) throw new Error('模板保存记录无效');
    return this.locked(async () => {
      let receipt = await this.#receipt(hash(`${hash(sessionId)}:${draftId}`));
      if (receipt?.status === 'pending') {
        const current = await this.readRecord(receipt.template.id);
        if (current.lastCommitId === receipt.commitId) await this.#reconcileRecord(current);
        receipt = await this.#receipt(receipt.commitId);
      }
      if (receipt?.status === 'committed') return receipt.template;
      const error = new Error('模板保存回执不存在'); error.code = 'ENOENT'; throw error;
    });
  }
  async previewSavedDraft(sessionId, draftId, page) {
    const record = await this.savedDraft(sessionId, draftId);
    return this.locked(async () => this.readPreview(await this.#recordPayload(record), record, page));
  }
  async cancel(sessionId, draftId) {
    return this.locked(() => this.audit(sessionId, 'cancel-personal-template', async () => {
      await rm(await this.draft(sessionId, draftId), { recursive: true });
      return { draftId };
    }));
  }
  async rename(sessionId, id, name) {
    return this.update(sessionId, id, { name }, 'rename-personal-template');
  }
  async update(sessionId, id, changes, operation = 'update-personal-template') {
    return this.locked(() => this.audit(sessionId, operation, async () => {
      const name = nameOf(changes.name);
      const description = changes.description === undefined ? undefined : descriptionOf(changes.description);
      const record = await this.readRecord(id);
      await this.#reconcileRecord(record);
      const next = { ...record, version: randomUUID(), name, previewTitle: name, ...(description === undefined ? {} : { description }), updatedAt: new Date().toISOString() };
      const file = path.join(this.root, 'saved', id, 'template.json');
      const temporary = `${file}.${randomUUID()}.tmp`;
      try {
        await writeFile(temporary, JSON.stringify(next), { mode: 0o600 });
        await this.catalogLocked(() => rename(temporary, file));
      } finally { await rm(temporary, { force: true }); }
      return next;
    }));
  }
  async remove(sessionId, id) {
    return this.locked(() => this.audit(sessionId, 'delete-personal-template', async () => {
      await this.#reconcileRecord(await this.readRecord(id));
      const directory = path.join(this.root, 'saved', id);
      const trash = path.join(await this.directory('trash'), `${id}-${randomUUID()}`);
      await this.catalogLocked(() => rename(directory, trash));
      return { id };
    }));
  }
  async #loadOwnedProject(directory) {
    // Walk before reading or copying; saved image paths remain confined to this project.
    const inspect = async current => {
      for (const item of await readdir(current, { withFileTypes: true })) {
        if (item.isSymbolicLink()) throw new Error('模板工程应使用本地文件');
        if (item.isDirectory()) await inspect(path.join(current, item.name));
        else if (!item.isFile()) throw new Error('模板工程包含无效文件');
      }
    };
    if ((await lstat(directory)).isSymbolicLink()) throw new Error('模板工程目录无效');
    await inspect(directory);
    const project = await loadPptdProject(directory);
    const checked = checkPptdProject(project);
    if (checked.status === 'fail') throw new Error(`模板工程校验失败：\n${checked.issues.filter(issue => issue.severity === 'error').map(issue => `${issue.file ?? ''} ${issue.message}`).join('\n')}`);
    return project;
  }
  async projectSource(id) {
    if (!ID.test(id)) throw new Error('个人模板标识无效');
    return this.catalogLocked(async () => {
      const directory = path.join(this.root, 'saved', id);
      const record = await this.#readDirectoryRecord(directory, id, true);
      return path.join(await this.#payloadDirectory(directory, record), 'project');
    });
  }
  async readSnapshot(id) {
    if (!ID.test(id)) throw new Error('个人模板标识无效');
    return this.catalogLocked(async () => {
      const directory = path.join(this.root, 'saved', id);
      const record = await this.#readDirectoryRecord(directory, id, true);
      const projectPath = path.join(await this.#payloadDirectory(directory, record), 'project');
      const design = (await readTemplateDesign(projectPath, record.pageIndex))?.design;
      return { record, projectPath, design };
    });
  }
  async readDesign(id) {
    if (!ID.test(id)) throw new Error('个人模板标识无效');
    return this.catalogLocked(async () => {
      const directory = path.join(this.root, 'saved', id);
      const record = await this.#readDirectoryRecord(directory, id, true);
      return (await readTemplateDesign(path.join(await this.#payloadDirectory(directory, record), 'project'), record.pageIndex))?.design;
    });
  }
  checkout(base, output, signal) {
    if (!ID.test(base?.templateId) || !VERSION.test(base?.version)) throw new Error('模板编辑来源无效');
    return this.locked(async () => {
      signal?.throwIfAborted();
      const record = await this.readRecord(base.templateId);
      if (record.version !== base.version) throw conflict();
      const source = await this.projectSource(record.id);
      const info = await lstat(output);
      if (!info.isDirectory() || info.isSymbolicLink() || (await readdir(output)).length) throw new Error('请使用新的空工程目录');
      for (const item of await readdir(source)) {
        signal?.throwIfAborted();
        await cp(path.join(source, item), path.join(output, item), { recursive: true, errorOnExist: true, force: false });
      }
      signal?.throwIfAborted();
      return record;
    });
  }
  async copyProject(id, output) {
    return this.locked(async () => {
      const source = await this.projectSource(id);
      for (const item of await readdir(source))
        await cp(path.join(source, item), path.join(output, item), { recursive: true, errorOnExist: true, force: false });
    });
  }
}
