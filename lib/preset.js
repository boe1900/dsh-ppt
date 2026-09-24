import { fileURLToPath } from "node:url";
import { PPT_SERVICE, PPT_SKILL_ROOT, registerPptSkill, registerPptTools } from "./index.js";

/** Agent-preset child plugin for DSH 0.1.7. */
export const name = "dsh-ppt-preset";
export const inject = ["tools", "systemPrompt", "skills", PPT_SERVICE, PPT_SKILL_ROOT];

export function apply(ctx) {
	const service = ctx.get(PPT_SERVICE) ?? ctx.root.get(PPT_SERVICE);
	if (service === void 0) throw new Error("dsh-ppt preset requires the host dsh-ppt plugin");
	const skillRoot = ctx.get(PPT_SKILL_ROOT) ?? ctx.root.get(PPT_SKILL_ROOT) ?? fileURLToPath(new URL("../skills/dsh-ppt", import.meta.url));
	registerPptSkill(ctx, skillRoot);
	registerPptTools(ctx, service);
}
