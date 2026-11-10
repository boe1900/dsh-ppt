/** Session-confined persistence for the DSH PPTD route. */
import type { SessionId } from '@deepseek-ai/dsh-session/types';
import type { PersonalTemplateLibrary } from './personal-templates.js';
import type { OfficeActivity, OfficePptState, OfficeTemplate } from './protocol.ts';
/** Storage ceilings used by the PPT route. */
export interface KimiStoreLimits {
    readonly maxDecksPerSession: number;
    readonly maxActivities: number;
}
/** DSH PPTD-compatible source published beside the editable PPTX. */
export interface PptdProjectFiles {
    readonly pptdEntryName: string;
    readonly pptdManifest: string;
    readonly pptdPages: readonly {
        readonly path: string;
        readonly content: string;
    }[];
    readonly assets: readonly {
        readonly path: string;
        readonly bytes: Uint8Array;
    }[];
}
/** Local store containing only PPT route state, source projects, and output files. */
export declare class PptStore {
    private readonly limits;
    readonly root: string;
    personalTemplates?: PersonalTemplateLibrary;
    private auditTail;
    constructor(root: string, limits: KimiStoreLimits, builtInTemplates: readonly OfficeTemplate[]);
    private sessionDirectory;
    private statePath;
    catalog(): Promise<OfficePptState>;
    readState(sessionId: SessionId): Promise<OfficePptState>;
    assertDeckCapacity(state: OfficePptState): void;
    writeState(state: OfficePptState): Promise<void>;
    writeJson(target: string, value: unknown): Promise<void>;
    recordAudit(sessionId: SessionId, activity: OfficeActivity, facts?: Readonly<Record<string, unknown>>): Promise<boolean>;
    prepareAudit(sessionId: SessionId, activity: OfficeActivity, facts?: Readonly<Record<string, unknown>>): Promise<void>;
    discardAudit(activity: OfficeActivity): Promise<void>;
    flushAudits(): Promise<void>;
    appendAudit(sessionId: SessionId, activity: OfficeActivity, facts: {
        readonly deckId?: string;
        readonly templateId?: string;
        readonly mode?: string;
    }): Promise<void>;
    private writeFile;
    private writeProjectFile;
}