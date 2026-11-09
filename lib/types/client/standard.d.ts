/** Native PPT preset chooser registration. */
import type { ClientContext, SessionId } from '@deepseek-ai/dsh-client-runtime/client';
import { type OfficePptHeroInjected } from './OfficePptHero.tsx';
/** Browser services required by the standard Composer adapter. */
export declare const standardInject: string[];
/** Build the session-scoped browser face shared by both Composer layouts. */
export declare function officePptHeroInjection(ctx: ClientContext): (sessionId: SessionId) => OfficePptHeroInjected;
/**
 * Register the PPT chooser in conversation.input.dock.
 * @param ctx - Browser context that owns the standard Composer slots.
 */
export declare function applyStandard(ctx: ClientContext): void;
