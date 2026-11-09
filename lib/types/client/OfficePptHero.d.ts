/** DSH PPT template chooser integrated into the blank-session composer. */
import type { SessionId } from '@deepseek-ai/dsh-client-runtime/client';
import type { PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots';
import { type OfficePptState, type OfficeTemplate } from '../protocol.ts';
import type { OfficePptClient } from './rpc.ts';
interface HeroState {
    readonly activeMode: 'ppt' | null;
    readonly pickerOpen: boolean;
    readonly loading: boolean;
    readonly templates: readonly OfficeTemplate[];
    readonly selectedId: string | null;
    readonly notice: string | false | null;
    readonly error: string;
}
/** Session-keyed presentation state shared by the chooser and personal template manager. */
export declare class OfficePptHeroStore {
    private readonly states;
    private readonly listeners;
    snapshot(sessionId: SessionId): HeroState;
    subscribe(sessionId: SessionId, listener: () => void): () => void;
    setPickerOpen(sessionId: SessionId, pickerOpen: boolean): void;
    setMode(sessionId: SessionId, activeMode: HeroState['activeMode']): void;
    setLoading(sessionId: SessionId, loading: boolean): void;
    setNotice(sessionId: SessionId, notice: HeroState['notice']): void;
    setError(sessionId: SessionId, error: string): void;
    setTemplates(sessionId: SessionId, templates: readonly OfficeTemplate[]): void;
    setTemplateState(sessionId: SessionId, next: OfficePptState): void;
    select(sessionId: SessionId, template: OfficeTemplate, activeMode: Exclude<HeroState['activeMode'], null>): void;
    deselect(sessionId: SessionId, activeMode: Exclude<HeroState['activeMode'], null>): void;
    private update;
}
/** Faces supplied by the package registration. */
export interface OfficePptHeroInjected {
    readonly client: OfficePptClient;
    readonly mode: OfficePptHeroStore;
}
/** Template chooser shown below the input only in blank native PPT sessions. */
export declare function OfficePptComposer(props: PropsRuntime<'conversation.input.dock'> & OfficePptHeroInjected & PropsLocale<'dsh-ppt'>): import("react").JSX.Element | null;
export {};
