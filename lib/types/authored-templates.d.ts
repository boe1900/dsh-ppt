import { PersonalTemplateLibrary, type PersonalTemplateRecord } from './personal-templates.js';
export interface AuthoredTemplateInput {
    readonly directory: string;
    readonly projectPath: string;
    readonly name: string;
    readonly description?: string;
    readonly turnId?: string;
}
export declare function templateName(value: unknown): string;
export declare function templateDescription(value?: unknown): string;
/** Snapshots only referenced native source, preserving reviewed examples and assets. */
export declare class AuthoredTemplateLibrary extends PersonalTemplateLibrary {
    prepareAuthored(sessionId: string, input: AuthoredTemplateInput, signal?: AbortSignal): Promise<{ draftId: string; template: PersonalTemplateRecord }>;
}
