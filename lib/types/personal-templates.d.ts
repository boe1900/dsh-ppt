import type { OfficeTemplate } from './protocol.js';
import type { PptStore } from './ppt-store.js';
import type { convertPptxToPptd } from './pptd-convert.js';
import type { writePptdProjectSource } from './pptd-publish.js';
import type { PersonalTemplateImportProgress } from './personal-template-import.js';
import type { PersonalTemplateUploadInput, PersonalTemplateUploadChunk, PersonalTemplateUploads } from './personal-template-upload.js';
import type { TemplateDesign } from './template-design.js';

export declare const PERSONAL_TEMPLATE_CONVERSION_VERSION = 10;
export declare const PERSONAL_TEMPLATE_PREVIEW_VERSION = 11;
export declare const MAX_PERSONAL_TEMPLATE_SLIDES = 200;
export declare const MAX_PERSONAL_TEMPLATE_HTTP_BODY_BYTES: number;
export type PersonalTemplateRecord = OfficeTemplate & {
    readonly origin: 'personal';
    readonly slideCount: number;
    readonly createdAt: string;
    readonly conversionVersion: number;
    readonly previewVersion: number;
    readonly pageIndex: readonly { slideNumber: number; file: string }[];
};
export type PersonalTemplatePrepareInput = (
    { readonly uploadId: string } | { readonly fileName: string; readonly base64: string }
) & { readonly requestId?: string; readonly draftId?: string };
export type PersonalTemplatePrepareResult =
    | { readonly duplicate: true; readonly template: PersonalTemplateRecord }
    | { readonly draftId: string; readonly template: PersonalTemplateRecord; readonly preview: string };
export interface PersonalTemplatePagePreview { readonly page: number; readonly preview: string; }
export interface PersonalTemplateProgress {
    readonly total: number;
    readonly completed: number;
    readonly pages: readonly PersonalTemplatePagePreview[];
    readonly status: 'waiting' | 'processing' | 'cancelled' | 'finished';
}
/** Shared host library. Uploads and drafts remain owned by their creating session. */
export declare class PersonalTemplateLibrary {
    readonly root: string;
    readonly maxSlides: number;
    readonly uploads: PersonalTemplateUploads;
    constructor(store: PptStore, convert: typeof convertPptxToPptd, writeSource: typeof writePptdProjectSource, maxSlides?: number);
    locked<T>(work: () => Promise<T>): Promise<T>;
    directory(...parts: string[]): Promise<string>;
    checkedFile(directory: string, file: string): Promise<string>;
    readRecord(id: string): Promise<PersonalTemplateRecord>;
    readDesign(id: string): Promise<TemplateDesign | undefined>;
    savedDraft(sessionId: string, draftId: string, id: string): Promise<PersonalTemplateRecord>;
    catalogLocked<T>(work: () => Promise<T>): Promise<T>;
    list(): Promise<PersonalTemplateRecord[]>;
    uploadStart(sessionId: string, input: PersonalTemplateUploadInput): ReturnType<PersonalTemplateUploads['start']>;
    uploadChunk(sessionId: string, input: PersonalTemplateUploadChunk): ReturnType<PersonalTemplateUploads['append']>;
    cancelUpload(sessionId: string, uploadId: string): Promise<{ uploadId: string }>;
    prepare(sessionId: string, input: PersonalTemplatePrepareInput): Promise<PersonalTemplatePrepareResult>;
    prepareSource(sessionId: string, input: PersonalTemplatePrepareInput, progress?: (progress: PersonalTemplateImportProgress) => void, signal?: AbortSignal): Promise<PersonalTemplatePrepareResult>;
    importProgress(sessionId: string, requestId: string, after?: number): Promise<PersonalTemplateProgress>;
    draft(sessionId: string, draftId: string): Promise<string>;
    previewPage(sessionId: string, draftId: string, page: number): Promise<PersonalTemplatePagePreview>;
    previewSavedPage(sessionId: string, templateId: string, page: number): Promise<PersonalTemplatePagePreview>;
    importPackage(sessionId: string, input: { fileName: string; base64: string }, signal?: AbortSignal): Promise<{ outcome: 'imported' | 'existing'; template: PersonalTemplateRecord }>;
    exportPackage(sessionId: string, templateId: string): Promise<{ fileName: string; base64: string }>;
    save(sessionId: string, draftId: string, name: string): Promise<PersonalTemplateRecord>;
    cancel(sessionId: string, draftId: string): Promise<{ draftId: string }>;
    rename(sessionId: string, id: string, name: string): Promise<PersonalTemplateRecord>;
    update(sessionId: string, id: string, changes: { name: string; description?: string }, operation?: string): Promise<PersonalTemplateRecord>;
    remove(sessionId: string, id: string): Promise<{ id: string }>;
    projectSource(id: string): Promise<string>;
    copyProject(id: string, output: string): Promise<void>;
}
