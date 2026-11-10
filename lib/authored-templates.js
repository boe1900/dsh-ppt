import { createHash, randomUUID } from 'node:crypto';
import { lstat, mkdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import sharp from 'sharp';
import { PersonalTemplateLibrary, PERSONAL_TEMPLATE_CONVERSION_VERSION, PERSONAL_TEMPLATE_PREVIEW_VERSION } from './personal-templates.js';
import { loadPptdProject, checkPptdProject } from './pptd.js';
import { extractPersonalTemplateProfile, validPersonalTemplateProfile } from './personal-template-profile.js';
import { renderOfficeTemplatePreview } from './office-template-preview.js';

export function templateName(value) {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > 80 || /[\u0000-\u001f\u007f]/u.test(value))
    throw new Error('模板名称应为 1–80 个字符');
  return value.trim();
}
export function templateDescription(value = '') {
  if (typeof value !== 'string' || value.length > 1000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value))
    throw new Error('模板描述应为 1000 个字符以内的文字');
  return value.trim();
}

/** Native templates share saved records and previews with imported templates. */
export class AuthoredTemplateLibrary extends PersonalTemplateLibrary {
  async prepareAuthored(sessionId, input, signal) {
    const name = templateName(input.name), description = templateDescription(input.description);
    return this.locked(() => this.audit(sessionId, 'prepare-authored-template', async () => {
      signal?.throwIfAborted();
      const draftId = randomUUID();
      const directory = await this.directory('drafts', draftId);
      try {
        const destination = path.join(directory, 'project');
        await mkdir(destination, { mode: 0o700 });
        const project = await loadPptdProject(input.directory);
        const source = project.source;
        if (source.entryName !== 'deck.pptd') throw new Error('请使用 deck.pptd 作为模板工程清单');
        const digest = createHash('sha256').update('dsh-native-template-v1\0');
        const files = [
          [source.entryName, Buffer.from(source.manifest)],
          ...[...source.pages].map(([file, text]) => [file, Buffer.from(text)]),
          ...[...source.assets.values()].map(asset => [asset.path, Buffer.from(asset.bytes)]),
        ].sort(([a], [b]) => a.localeCompare(b));
        if (files.length > 1000 || files.reduce((sum, [, bytes]) => sum + bytes.length, 0) > 128 * 1024 * 1024)
          throw new Error('模板工程超过文件或资源大小限制');
        for (const [relative, bytes] of files) {
          signal?.throwIfAborted();
          let current = input.directory;
          for (const part of relative.split('/')) {
            current = path.join(current, part);
            const info = await lstat(current);
            if (info.isSymbolicLink()) throw new Error(`模板工程不能包含符号链接：${relative}`);
          }
          digest.update(JSON.stringify([relative, bytes.length])).update('\0').update(bytes);
        }
        await this.writeSource(destination, source);
        if (!project.pages.length || project.pages.length > this.maxSlides) throw new Error(`模板应包含 1–${this.maxSlides} 页`);
        if (project.pages.some(page => !/^pages\/[A-Za-z0-9_.-]+\.page$/.test(page.file)))
          throw new Error('请将模板页面放在 pages/ 下并使用字母、数字、下划线或连字符命名，例如 pages/cover.page');
        const checked = checkPptdProject(project);
        if (checked.status === 'fail') throw new Error(`请先修复模板排版：\n${checked.issues.filter(item => item.severity === 'error').map(item => `${item.file ?? ''} ${item.message}`).join('\n')}`);
        const sha256 = digest.digest('hex'), id = `personal-${sha256}`;
        const profile = extractPersonalTemplateProfile(project);
        const pageIndex = project.pages.map((page, i) => ({ slideNumber: i + 1, file: page.file }));
        if (!validPersonalTemplateProfile(profile, pageIndex)) throw new Error('模板结构索引生成失败，请检查模板工程');
        await this.ensureFontAccess();
        const previewMetadata = await renderOfficeTemplatePreview(project, directory, 1920, this.maxSlides, undefined, signal);
        signal?.throwIfAborted();
        const representative = [1, Math.ceil(project.pages.length / 2), project.pages.length];
        const previewImages = await Promise.all(representative.map(async page => {
          const bytes = await sharp(path.join(directory, 'preview', 'pages', `page-${page}.png`))
            .resize({ width: 720, height: 720, fit: 'inside', withoutEnlargement: true }).png().toBuffer();
          return `data:image/png;base64,${bytes.toString('base64')}`;
        }));
        const record = {
          id, conversionVersion: PERSONAL_TEMPLATE_CONVERSION_VERSION, previewVersion: PERSONAL_TEMPLATE_PREVIEW_VERSION,
          origin: 'personal', authoringKind: 'native-pptd', category: 'personal', supportedModes: ['ppt'], name, description,
          previewTitle: name, previewSubtitle: '我的模板', createdAt: new Date().toISOString(), fileName: 'deck.pptd', sha256,
          slideCount: project.pages.length, aspectRatio: project.width / project.height >= 1.5 ? 'wide' : 'standard',
          templateProfile: profile, palette: profile.palette,
          titleFontFace: profile.typography.title.fontFace, bodyFontFace: profile.typography.body.fontFace,
          previewImages, previewMetadata, diagnostics: [], resourcePlaceholders: [], checkWarnings: checked.warningCount,
          reviewIssues: checked.issues.filter(item => item.severity === 'warning'),
          pageIndex,
        };
        await writeFile(path.join(directory, 'template.json'), JSON.stringify(record), { mode: 0o600 });
        await writeFile(path.join(directory, 'owner.json'), JSON.stringify({ session: createHash('sha256').update(sessionId).digest('hex') }), { mode: 0o600 });
        signal?.throwIfAborted();
        return { draftId, template: record };
      } catch (error) {
        await rm(directory, { recursive: true, force: true });
        throw error;
      }
    }));
  }
}
