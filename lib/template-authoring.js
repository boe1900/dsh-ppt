import { randomUUID } from 'node:crypto';
import { templateName } from './authored-templates.js';

const uuid = value => typeof value === 'string' && /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(value);
const templateId = value => typeof value === 'string' && /^personal-[a-f0-9]{64}$/.test(value);
const projectPath = value => typeof value === 'string' && value.length < 4096 && value !== '' && !value.startsWith('/')
  && !value.includes('\\') && !value.split('/').some(part => part === '..');
export function normalizeTemplateAuthoring(value) {
  if (!value || !uuid(value.workflowId) || !Number.isSafeInteger(value.revision) || value.revision < 0) return undefined;
  const base = { workflowId: value.workflowId, revision: value.revision, guideSent: value.guideSent === true };
  if (typeof value.turnId === 'string' && value.turnId.length <= 128) base.turnId = value.turnId;
  if (projectPath(value.projectPath)) base.projectPath = value.projectPath;
  const named = typeof value.name === 'string' && value.name.trim().length > 0 && value.name.length <= 80 && !/[\u0000-\u001f\u007f]/u.test(value.name);
  if (named) base.name = value.name;
  if (value.phase === 'creating') return { ...base, phase: 'creating' };
  if (value.phase === 'review' && named && uuid(value.draftId) && templateId(value.templateId) && projectPath(value.projectPath)
    && Number.isInteger(value.slideCount) && value.slideCount > 0 && value.slideCount <= 200)
    return { ...base, phase: 'review', draftId: value.draftId, templateId: value.templateId, slideCount: value.slideCount };
  if (value.phase === 'saved' && named && templateId(value.templateId)) return { ...base, phase: 'saved', templateId: value.templateId };
}

/** The session service owns the workflow; render jobs never own session state. */
export class TemplateAuthoring {
  jobs = new Map();
  constructor(service) { this.service = service; }
  change(sessionId, operation, work) {
    return this.service.mutate(sessionId, operation, { kind: 'user' }, async state => {
      const value = await work(state.templateAuthoring, state);
      return { state: { ...state, templateAuthoring: value }, value, summary: operation };
    });
  }
  async discard(sessionId, draftId) {
    if (draftId) await this.service.store.personalTemplates.cancel(sessionId, draftId).catch(error => {
      if (error.code !== 'ENOENT') throw error;
    });
  }
  start(sessionId) {
    return this.change(sessionId, 'start-template-authoring', current => current && current.phase !== 'saved' ? current
      : { phase: 'creating', workflowId: randomUUID(), revision: 0, guideSent: false });
  }
  guided(sessionId, workflowId) {
    return this.change(sessionId, 'guide-template-authoring', current => {
      if (!current || current.workflowId !== workflowId) throw new Error('模板制作流程已结束，请重新开始');
      return { ...current, guideSent: true };
    });
  }
  async exit(sessionId, workflowId) {
    await this.change(sessionId, 'exit-template-authoring', async current => {
      if (!current || current.workflowId !== workflowId) return current;
      this.jobs.get(sessionId)?.abort();
      await this.discard(sessionId, current.draftId);
      return undefined;
    });
  }
  async revise(sessionId, turnId) {
    return this.change(sessionId, 'revise-template-authoring', async current => {
      if (current?.phase !== 'review' || !turnId || current.turnId === turnId) return current;
      await this.discard(sessionId, current.draftId);
      const { draftId, templateId, slideCount, ...base } = current;
      return { ...base, phase: 'creating', revision: current.revision + 1 };
    });
  }
  async prepare(sessionId, input, signal) {
    const controller = new AbortController();
    let reservation, previousDraft;
    await this.change(sessionId, 'prepare-template-authoring', current => {
      if (!current || current.phase === 'saved') throw new Error('请先在模板面板点击「新建自定义模板」');
      this.jobs.get(sessionId)?.abort();
      this.jobs.set(sessionId, controller);
      previousDraft = current.draftId;
      reservation = { ...current, phase: 'creating', revision: current.revision + 1,
        projectPath: input.projectPath, name: input.name, turnId: input.turnId, guideSent: true };
      delete reservation.draftId; delete reservation.templateId; delete reservation.slideCount;
      return reservation;
    });
    let prepared;
    try {
      prepared = await this.service.store.personalTemplates.prepareAuthored(sessionId, input,
        signal ? AbortSignal.any([signal, controller.signal]) : controller.signal);
      return await this.change(sessionId, 'review-template-authoring', current => {
        if (!current || current.workflowId !== reservation.workflowId || current.revision !== reservation.revision)
          throw new Error('模板制作流程已更新，请使用最新预览');
        controller.signal.throwIfAborted();
        return { ...reservation, phase: 'review', draftId: prepared.draftId, templateId: prepared.template.id,
          slideCount: prepared.template.slideCount };
      });
    } catch (error) {
      await this.discard(sessionId, prepared?.draftId);
      throw error;
    } finally {
      if (this.jobs.get(sessionId) === controller) this.jobs.delete(sessionId);
      await this.discard(sessionId, previousDraft);
    }
  }
  matches(current, request) {
    if (!current || current.workflowId !== request.workflowId || current.revision !== request.revision
      || (current.phase === 'review' && current.draftId !== request.draftId))
      throw new Error('预览版本已更新，请检查最新预览后再保存');
  }
  preview(sessionId, request) {
    return this.service.withLock(sessionId, async () => {
      const current = (await this.service.state(sessionId)).templateAuthoring;
      this.matches(current, request);
      if (current.phase !== 'review') throw new Error('当前没有待保存的模板预览');
      try { return await this.service.store.personalTemplates.previewPage(sessionId, current.draftId, request.page); }
      catch (error) {
        if (error.code !== 'ENOENT') throw error;
        return this.service.store.personalTemplates.previewSavedPage(sessionId, current.templateId, request.page);
      }
    });
  }
  async save(sessionId, request) {
    const name = templateName(request.name);
    return this.change(sessionId, 'save-authored-template', async current => {
      this.matches(current, request);
      if (current.phase === 'saved') return current;
      if (current.phase !== 'review') throw new Error('请先生成并预览模板');
      let saved;
      try { saved = await this.service.store.personalTemplates.save(sessionId, current.draftId, name); }
      catch (error) {
        if (error.code !== 'ENOENT') throw error;
        saved = await this.service.store.personalTemplates.readRecord(current.templateId);
      }
      return { phase: 'saved', workflowId: current.workflowId, revision: current.revision,
        templateId: saved.id, name: saved.name, guideSent: true };
    });
  }
}
