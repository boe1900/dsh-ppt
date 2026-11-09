/** Typed browser wrapper over the dedicated Office PPT Connection channel. */
import type { ClientConnectionRpc } from '@deepseek-ai/dsh-client-connection/client';
import type { SessionId } from '@deepseek-ai/dsh-client-runtime/client';
/** Session-bound Office PPT client used by the view. */
export interface OfficePptClient {
    readonly bound: boolean;
    call<T>(endpoint: string, payload?: Record<string, unknown>, signal?: AbortSignal): Promise<T>;
}
/**
 * Bind every request to the active conversation session.
 * @param rpc - Browser Connection RPC client.
 * @param sessionId - Active conversation session.
 * @returns Session-bound Office PPT client.
 */
export declare function createOfficePptClient(rpc: ClientConnectionRpc, sessionId: SessionId): OfficePptClient;
