import { randomUUID } from 'node:crypto';
import { templateName, templateDescription } from './authored-templates.js';

const uuid = value => typeof value === 'string' && /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(value);
const templateId = value => typeof value === 'string' && /^personal-[a-f0-9]{64}$/.test(value);
const version = value => typeof value === 'string' && /^(?:[a-f0-9]{64}|[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12})$/.test(value);
const validBase = value => templateId(value?.templateId) && version(value?.version);
const baseOf = record => ({ templateId: record.id, version: record.version });
const projectPath = value => typeof value === 'string' && value.length < 4096 && value !== '' && !value.startsWith('/')
  && !value.includes('\\') && !value.split('/').some(part => part === '..');
export function normalizeTemplateAuthoring(value) {
  if (!value || !uuid(value.workflowId) || !Number.isSafeInteger(value.revision) || value.revision < 0) return undefined;
  if (value.origin !== undefined && value.origin?.kind !== 'new' && !(value.origin?.kind === 'existing' && validBase(value.origin.base))) return undefined;
  const origin = value.origin?.kind === 'existing' ? { kind: 'existing', base: { templateId: value.origin.base.templateId, version: value.origin.base.version } } : { kind: 'new' };
  const base = { workflowId: value.workflowId, revision: value.revision, guideSent: value.guideSent === true, origin };
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
  if (value.phase === 'saved' && named && templateId(value.templateId)) {
    let lastSave;
    const saved = value.lastSave;
    if (saved !== undefined) {
      if (!uuid(saved.draftId) || typeof saved.name !== 'string' || !saved.name.trim() || saved.name.length > 80
        || !['copy', 'original'].includes(saved.destination?.kind) || (saved.destination.kind === 'original' && !validBase(saved.destination.base))) return undefined;
      lastSave = { draftId: saved.draftId, name: saved.name, destination: saved.destination.kind === 'copy' ? { kind: 'copy' }
        : { kind: 'original', base: { templateId: saved.destination.base.templateId, version: saved.destination.base.version } } };
    }
    return { ...base, phase: 'saved', templateId: value.templateId, ...(lastSave ? { lastSave } : {}) };
  }
}

export class TemplateAuthoring {
  jobs = new Map();
  constructor(service) { this.service = service; }
  project(state) {
    let current = state.templateAuthoring;
    if (current?.phase === 'saved' && current.origin.kind === 'new') {
      const record = state.templates.find(record => record.id === current.templateId && version(record.version));
      if (record) { current = { ...current, origin: { kind: 'existing', base: baseOf(record) } }; state = { ...state, templateAuthoring: current }; }
    }
    if (current?.phase !== 'preparing' || this.jobs.has(state.sessionId)) return state;
    return { ...state, templateAuthoring: { ...current, phase: 'failed', error: '预览生成已中断，请重试；工程文件已保留。' } };
  }
  change(sessionId, operation, work) {
    return this.service.mutate(sessionId, operation, { kind: 'user' }, async state => {
      state = this.project(state);
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
  start(sessionId, id) {
    if (id !== undefined && !templateId(id)) throw new Error('请选择个人模板继续编辑');
    return this.change(sessionId, 'start-template-authoring', async current => {
      if (current && current.phase !== 'saved') {
        if ((!id && current.origin.kind === 'new') || (id && current.origin.kind === 'existing' && current.origin.base.templateId === id)) return current;
        throw new Error('请先继续或退出当前模板制作流程');
      }
      const record = id ? await this.service.store.personalTemplates.readRecord(id) : undefined;
      return { phase: 'creating', workflowId: randomUUID(), revision: 0, guideSent: false,
        origin: record ? { kind: 'existing', base: baseOf(record) } : { kind: 'new' },
        ...(record ? { name: record.name, description: record.description } : {}) };
    });
  }
  copySource(sessionId, input, publishCopy, signal) {
    if (!templateId(input.templateId) || !projectPath(input.projectPath)) throw new Error('模板编辑工程参数无效');
    return this.change(sessionId, 'copy-template-authoring-source', async current => {
      signal?.throwIfAborted();
      const matching = current?.origin.kind === 'existing' && current.origin.base.templateId === input.templateId;
      if (current && current.phase !== 'saved' && !matching) throw new Error('请先继续或退出当前模板制作流程');
      if (matching && current.projectPath) return current;
      const selected = matching ? undefined : await this.service.store.personalTemplates.readRecord(input.templateId);
      const origin = matching ? current.origin : { kind: 'existing', base: baseOf(selected) };
      const record = await publishCopy(origin.base);
      signal?.throwIfAborted();
      if (record.id !== origin.base.templateId || record.version !== origin.base.version) throw new Error('模板工程副本与编辑来源不一致');
      return { phase: 'creating', workflowId: matching ? current.workflowId : randomUUID(), revision: matching ? current.revision : 0,
        origin, guideSent: true, projectPath: input.projectPath, name: record.name, description: record.description,
        ...(input.turnId ? { turnId: input.turnId } : {}) };
    });
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
    const next = await this.change(sessionId, 'revise-template-authoring', async current => {
      if (!['review', 'saved'].includes(current?.phase) || !turnId || current.turnId === turnId) return current;
      if (current.phase === 'saved' && current.origin.kind === 'new') return current;
      obsolete = current;
      const { draftId, templateId, slideCount, lastSave, ...base } = current;
      return { ...base, phase: 'creating', revision: current.revision + 1, turnId };
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
        const base = current && !(current.phase === 'saved' && current.projectPath !== input.projectPath) ? current
          : { workflowId: randomUUID(), revision: 0, origin: { kind: 'new' } };
        if (base.origin.kind === 'existing' && base.projectPath && base.projectPath !== input.projectPath) throw new Error('当前正在编辑另一模板工程，请先退出模板制作流程');
        previousDraft = base.draftId;
        reservation = { phase: 'preparing', workflowId: base.workflowId, revision: base.revision + 1,
          projectPath: input.projectPath, name, description, turnId: input.turnId, guideSent: true, origin: base.origin };
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
      || (current.phase === 'review' && current.draftId !== request.draftId)
      || (current.phase === 'saved' && current.lastSave?.draftId !== request.draftId))
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
        return this.service.store.personalTemplates.previewSavedDraft(sessionId, current.draftId, request.page);
      }
    });
  }
  async save(sessionId, request) {
    const name = templateName(request.name);
    return this.change(sessionId, 'save-authored-template', async current => {
      this.matches(current, request);
      if (!['review', 'saved'].includes(current.phase)) throw new Error('请先生成并预览模板');
      const kind = request.destination ?? (current.phase === 'saved' ? current.lastSave.destination.kind : current.origin.kind === 'existing' ? 'original' : 'copy');
      if (!['copy', 'original'].includes(kind)) throw new Error('模板保存目标无效');
      let destination;
      if (current.phase === 'saved') {
        if (kind !== current.lastSave.destination.kind) throw new Error('该预览已使用其他保存目标提交，请生成新的预览');
        destination = current.lastSave.destination;
      } else if (kind === 'original') {
        if (current.origin.kind !== 'existing') throw new Error('新模板只能另存为新模板');
        destination = { kind: 'original', base: current.origin.base };
      } else destination = { kind: 'copy' };
      const result = await this.service.store.personalTemplates.commit(sessionId, request.draftId, { name, destination });
      return { phase: 'saved', workflowId: current.workflowId, revision: current.revision, projectPath: current.projectPath,
        templateId: result.template.id, name: result.template.name, description: result.template.description, guideSent: true,
        turnId: current.turnId, origin: { kind: 'existing', base: result.base }, lastSave: { draftId: request.draftId, name, destination } };
    });
  }
}
