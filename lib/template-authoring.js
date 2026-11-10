import { randomUUID } from 'node:crypto';
import { templateName, templateDescription } from './authored-templates.js';

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
  if (typeof value.description === 'string' && value.description.length <= 1000) base.description = value.description;
  if (value.phase === 'creating') return { ...base, phase: 'creating' };
  if (['preparing', 'failed', 'cancelled'].includes(value.phase) && named && base.projectPath) return {
    ...base, phase: value.phase,
    ...typeof value.error === 'string' ? { error: value.error.slice(0, 2000) } : {},
  };
  if (value.phase === 'review' && named && uuid(value.draftId) && templateId(value.templateId) && projectPath(value.projectPath)
    && Number.isInteger(value.slideCount) && value.slideCount > 0 && value.slideCount <= 200)
    return { ...base, phase: 'review', draftId: value.draftId, templateId: value.templateId, slideCount: value.slideCount };
  if (value.phase === 'saved' && named && templateId(value.templateId)) return { ...base, phase: 'saved', templateId: value.templateId };
}

export class TemplateAuthoring {
  jobs = new Map();
  constructor(service) { this.service = service; }
  project(state) {
    const current = state.templateAuthoring;
    if (current?.phase !== 'preparing' || this.jobs.has(state.sessionId)) return state;
    return { ...state, templateAuthoring: { ...current, phase: 'failed', error: '预览生成已中断，请重试；工程文件已保留。' } };
  }
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
  async cleanup(sessionId, draftId) {
    try { await this.discard(sessionId, draftId); }
    catch (error) {
      await this.service.store.recordAudit(sessionId, { id: randomUUID(), operation: 'cleanup-template-draft',
        actor: 'user', status: 'failed', startedAt: new Date().toISOString(), completedAt: new Date().toISOString(), error: error.message }, { draftId });
    }
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
    let obsolete, job;
    await this.change(sessionId, 'exit-template-authoring', current => {
      if (!current || current.workflowId !== workflowId) return current;
      obsolete = current;
      job = this.jobs.get(sessionId);
      return undefined;
    });
    if (obsolete) {
      job?.abort();
      await this.cleanup(sessionId, obsolete.draftId);
    }
    return Boolean(obsolete);
  }
  async revise(sessionId, turnId) {
    let obsolete;
    const next = await this.change(sessionId, 'revise-template-authoring', current => {
      if (current?.phase !== 'review' || !turnId || current.turnId === turnId) return current;
      obsolete = current;
      const { draftId, templateId, slideCount, ...base } = current;
      return { ...base, phase: 'creating', revision: current.revision + 1 };
    });
    await this.cleanup(sessionId, obsolete?.draftId);
    return next;
  }
  async prepare(sessionId, input, signal) {
    const name = templateName(input.name), description = templateDescription(input.description);
    const controller = new AbortController();
    let previousJob;
    const prepareSignal = signal ? AbortSignal.any([signal, controller.signal]) : controller.signal;
    prepareSignal.throwIfAborted();
    let reservation, previousDraft, prepared, admitted = false;
    try {
      await this.change(sessionId, 'prepare-template-authoring', current => {
        prepareSignal.throwIfAborted();
        const base = current && current.phase !== 'saved' ? current : { workflowId: randomUUID(), revision: 0 };
        previousDraft = base.draftId;
        reservation = { phase: 'preparing', workflowId: base.workflowId, revision: base.revision + 1,
          projectPath: input.projectPath, name, description, turnId: input.turnId, guideSent: true };
        previousJob = this.jobs.get(sessionId);
        this.jobs.set(sessionId, controller);
        return reservation;
      });
      admitted = true;
      previousJob?.abort();
      prepareSignal.throwIfAborted();
      prepared = await this.service.store.personalTemplates.prepareAuthored(sessionId, input, prepareSignal);
      return await this.change(sessionId, 'review-template-authoring', current => {
        this.matches(current, reservation);
        prepareSignal.throwIfAborted();
        return { ...reservation, phase: 'review', draftId: prepared.draftId, templateId: prepared.template.id,
          slideCount: prepared.template.slideCount };
      });
    } catch (error) {
      await this.cleanup(sessionId, prepared?.draftId);
      if (admitted) await this.change(sessionId, 'fail-template-authoring', current => {
        if (!current || current.workflowId !== reservation.workflowId || current.revision !== reservation.revision) return current;
        return { ...reservation, phase: prepareSignal.aborted ? 'cancelled' : 'failed',
          error: prepareSignal.aborted ? '预览生成已取消；工程文件已保留，可重试或在对话中继续修改。' : String(error.message ?? error).slice(0, 2000) };
      }).catch(() => {});
      throw error;
    } finally {
      if (this.jobs.get(sessionId) === controller) {
        if (!admitted && previousJob) this.jobs.set(sessionId, previousJob);
        else this.jobs.delete(sessionId);
      }
      if (admitted) await this.cleanup(sessionId, previousDraft);
    }
  }
  matches(current, request) {
    if (!current || current.workflowId !== request.workflowId || current.revision !== request.revision
      || (current.phase === 'review' && current.draftId !== request.draftId))
      throw new Error('预览版本已更新，请检查最新预览后再保存');
  }
  async savedReceipt(sessionId, current) {
    return this.service.store.personalTemplates.savedDraft(sessionId, current.draftId, current.templateId);
  }
  preview(sessionId, request) {
    return this.service.withLock(sessionId, async () => {
      const current = (await this.service.state(sessionId)).templateAuthoring;
      this.matches(current, request);
      if (current.phase !== 'review') throw new Error('当前没有待保存的模板预览');
      try { return await this.service.store.personalTemplates.previewPage(sessionId, current.draftId, request.page); }
      catch (error) {
        if (error.code !== 'ENOENT') throw error;
        await this.savedReceipt(sessionId, current);
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
      try {
        saved = await this.service.store.personalTemplates.save(sessionId, current.draftId, name);
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
        saved = await this.savedReceipt(sessionId, current);
      }
      return { phase: 'saved', workflowId: current.workflowId, revision: current.revision,
        templateId: saved.id, name: saved.name, guideSent: true };
    });
  }
}
