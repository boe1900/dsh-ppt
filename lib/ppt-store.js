import { createHash, randomUUID } from 'node:crypto';
import { appendFile, mkdir, open, readFile, readdir, rename, rm } from 'node:fs/promises';
import path from 'node:path';
import { normalizeTemplateAuthoring } from './template-authoring.js';

//#region lib/types/ppt-store.js
/** Session-confined persistence for the DSH PPTD route. */
function sessionKey(sessionId) {
	return createHash("sha256").update(sessionId).digest("hex").slice(0, 32);
}
/** Retired built-ins never re-enter the live catalog through persisted state. */
function persistedState(value, sessionId, personalTemplates = [], builtInTemplates) {
 const record = typeof value === "object" && value !== null && !Array.isArray(value) ? value : {};
 const legacySelection = typeof record.selectedTemplateId === "string" ? record.selectedTemplateId : undefined;
 const requested = legacySelection?.replace(/^kimi-(work|consulting)-curated-/, "dsh-$1-curated-");
 const templates = [...builtInTemplates, ...personalTemplates];
 const selected = templates.find(template => template.id === requested);
 const missingPersonal = requested?.startsWith("personal-") && selected === undefined;
 const retired = requested !== undefined && selected === undefined && !missingPersonal;
 const fallback = builtInTemplates.find(template => template.id === "dsh-engineering-blueprint") ?? builtInTemplates[0];
 const selectedTemplateId = retired ? fallback.id : selected?.id;
 const templateAuthoring = normalizeTemplateAuthoring(record.templateAuthoring);
 return {
  sessionId, templates,
  ...(templateAuthoring ? { templateAuthoring } : {}),
  // Preserve historical decks and generated files; they are user-owned records.
  decks: Array.isArray(record.decks) ? record.decks : [],
  activities: Array.isArray(record.activities) ? record.activities : [],
  ...(['word', 'excel'].includes(record.documentMode) ? { documentMode: record.documentMode } :
      record.presentationMode === "ppt" ? { presentationMode: "ppt" } : {}),
  ...(['word', 'excel'].includes(record.documentMode) && record.selectedDocumentTemplate?.mode === record.documentMode
      && /^[a-z0-9-]{1,96}$/.test(record.selectedDocumentTemplate?.id ?? '') && /^[a-zA-Z0-9.-]{1,96}$/.test(record.selectedDocumentTemplate?.revision ?? '')
      ? { selectedDocumentTemplate: { id: record.selectedDocumentTemplate.id, mode: record.documentMode, revision: record.selectedDocumentTemplate.revision } } : {}),
  ...(selectedTemplateId === undefined ? {} : { selectedTemplateId }),
  ...(missingPersonal ? { templateMigration: { reason: "personal-template-deleted" } } : retired ? { templateMigration: { reason: "template-retired", replacementId: fallback.id } } :
     record.templateMigration?.reason === "template-retired" ? { templateMigration: record.templateMigration } : {})
 };
}
/** Local store containing only PPT route state, source projects, and output files. */
export class PptStore {
	limits;
	root;
	auditTail = Promise.resolve();
	constructor(root, limits, builtInTemplates) {
		this.builtInTemplates = builtInTemplates;
		this.limits = limits;
		if (!path.isAbsolute(root)) throw new Error("DSH PPT root must be an absolute path");
		this.root = path.resolve(root);
	}
	sessionDirectory(sessionId) {
		return path.join(this.root, "sessions", sessionKey(sessionId));
	}
	statePath(sessionId) {
		return path.join(this.sessionDirectory(sessionId), "state.json");
	}
	async catalog() {
		const personalTemplates = await this.personalTemplates?.list() ?? [];
		return persistedState({}, "", personalTemplates, this.builtInTemplates);
	}
	async readState(sessionId) {
		const personalTemplates = await this.personalTemplates?.list() ?? [];
		try {
			return persistedState(JSON.parse(await readFile(this.statePath(sessionId), "utf8")), sessionId, personalTemplates, this.builtInTemplates);
		} catch (error) {
			if (error.code !== "ENOENT") throw error;
			return {
				sessionId,
				templates: [...this.builtInTemplates, ...personalTemplates],
				decks: [],
				activities: []
			};
		}
	}
	assertDeckCapacity(state) {
		if (state.decks.length >= this.limits.maxDecksPerSession) throw new Error("session deck limit exceeded");
	}
	async writeState(state) {
		if (state.decks.length > this.limits.maxDecksPerSession) throw new Error("session deck limit exceeded");
		const bounded = {
			...state,
			templates: this.builtInTemplates,
			activities: state.activities.slice(0, this.limits.maxActivities)
		};
        await this.writeJson(this.statePath(state.sessionId), bounded);
    }
    async writeJson(target, value) {
        await mkdir(path.dirname(target), { recursive: true, mode: 0o700 });
        const temporary = `${target}.${randomUUID()}.tmp`;
        try {
            await this.writeFile(temporary, `${JSON.stringify(value, null, 2)}\n`, 0o600);
            await rename(temporary, target);
        } finally { await rm(temporary, { force: true }).catch(() => {}); }
    }
    /** Audit delivery is recoverable side work and never rolls back a committed business mutation. */
    async recordAudit(sessionId, activity, facts = {}) {
        const target = path.join(this.root, 'audit-pending', `${activity.id}.json`);
        const entry = { sessionId, activity, facts };
        try { await this.writeJson(target, entry); }
        catch { /* The session activity remains the committed record if the volume cannot accept an outbox file. */ }
        try {
            await this.appendAudit(sessionId, activity, facts);
            await rm(target, { force: true }).catch(() => {});
            return true;
        } catch { return false; }
    }
    async prepareAudit(sessionId, activity, facts = {}) {
        await this.writeJson(path.join(this.root, 'audit-pending', `${activity.id}.json`), { sessionId, activity, facts, requiresCommit: true });
    }
    async discardAudit(activity) {
        await rm(path.join(this.root, 'audit-pending', `${activity.id}.json`), { force: true }).catch(() => {});
    }
    async flushAudits() {
        const directory = path.join(this.root, 'audit-pending');
        let names;
        try { names = await readdir(directory); }
        catch (error) { if (error.code === 'ENOENT') return; throw error; }
        for (const name of names.filter(name => /^[a-f0-9-]{36}\.json$/.test(name))) {
            const target = path.join(directory, name);
            try {
                const { sessionId, activity, facts, requiresCommit } = JSON.parse(await readFile(target, 'utf8'));
                if (requiresCommit) {
                    const state = JSON.parse(await readFile(this.statePath(sessionId), 'utf8'));
                    if (!state.activities.some(item => item.id === activity.id)) continue;
                    await this.writeJson(target, { sessionId, activity, facts });
                }
                await this.appendAudit(sessionId, activity, facts);
                await rm(target, { force: true });
            } catch { /* Keep the entry for the next delivery attempt. Activity IDs identify duplicates after a crash. */ }
        }
    }

	async appendAudit(sessionId, activity, facts) {
		const line = JSON.stringify({
			sessionKey: sessionKey(sessionId),
			...activity,
			...facts
		});
		const write = this.auditTail.catch(() => void 0).then(async () => {
			await mkdir(this.root, {
				recursive: true,
				mode: 448
			});
			await appendFile(path.join(this.root, "audit.ndjson"), `${line}\n`, {
				encoding: "utf8",
				mode: 384
			});
		});
		this.auditTail = write;
		await write;
	}
	async writeFile(target, content, mode) {
		const handle = await open(target, "wx", mode);
		try {
			await handle.writeFile(content);
			await handle.sync();
		} finally {
			await handle.close();
		}
	}
	async writeProjectFile(root, relative, content) {
		if (relative === "" || path.isAbsolute(relative) || relative.includes("\\") || relative.includes("\0")) throw new Error(`DSH PPT project path is unsafe: ${relative}`);
		const normalized = path.posix.normalize(relative);
		if (normalized === "." || normalized === ".." || normalized.startsWith("../") || normalized.startsWith("/")) throw new Error(`DSH PPT project path is unsafe: ${relative}`);
		const target = path.join(root, ...normalized.split("/"));
		await mkdir(path.dirname(target), {
			recursive: true,
			mode: 493
		});
		await this.writeFile(target, content, 420);
	}
};
//#endregion
