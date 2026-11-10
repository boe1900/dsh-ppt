import type { TemplateAuthoringState } from './protocol.js';
import type { PersonalTemplatePagePreview } from './personal-templates.js';
import type { AuthoredTemplateInput } from './authored-templates.js';
import type { PptService } from './ppt-service.js';
export interface TemplateReviewRequest {
    readonly workflowId: string;
    readonly revision: number;
    readonly draftId: string;
}
export declare function normalizeTemplateAuthoring(value: unknown): TemplateAuthoringState | undefined;
/** Transitions are serialized through the owning session service. */
export declare class TemplateAuthoring {
    constructor(service: PptService);
    start(sessionId: string): Promise<TemplateAuthoringState>;
    guided(sessionId: string, workflowId: string): Promise<TemplateAuthoringState>;
    exit(sessionId: string, workflowId: string): Promise<void>;
    revise(sessionId: string, turnId?: string): Promise<TemplateAuthoringState | undefined>;
    prepare(sessionId: string, input: AuthoredTemplateInput, signal?: AbortSignal): Promise<Extract<TemplateAuthoringState, { phase: 'review' }>>;
    preview(sessionId: string, request: TemplateReviewRequest & { page: number }): Promise<PersonalTemplatePagePreview>;
    save(sessionId: string, request: TemplateReviewRequest & { name: string }): Promise<Extract<TemplateAuthoringState, { phase: 'saved' }>>;
}
