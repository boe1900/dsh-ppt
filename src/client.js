window.__ModuleLoader__.load({
	id: "@cola1900/dsh-ppt",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		let react = require("react");
		let react_jsx_runtime = require("react/jsx-runtime");
		//#region ../office-ppt/src/protocol.ts
		/**
		* Resolve workflow availability while preserving compatibility with older extracted templates.
		* @param template - Persisted or built-in template definition.
		* @param mode - Presentation workflow requesting the template.
		* @returns Whether the template is available to the requested workflow.
		*/
		function templateSupportsMode(template, mode) {
			return template.supportedModes?.includes(mode) ?? true;
		}
		//#endregion
		//#region ../office-ppt/src/client/curated-previews.ts
		/** Browser preview registry for every source-backed catalog template. */
		/** Content-addressed local image URLs; both clients share the core JPG files. */
		const CURATED_TEMPLATE_PREVIEWS = /* GENERATED_PPT_PREVIEWS */ {};
		//#endregion
		//#region \0dsh-css:ppt-composer.css.mjs
		const css = "._2S_x-q_modeRoot{box-sizing:border-box;width:100%;padding:0 var(--dsh-composer-side-clearance);position:relative}._2S_x-q_modeRow{flex-wrap:wrap;justify-content:flex-start;align-items:center;gap:8px;min-height:32px;display:flex}._2S_x-q_modeChip{border:1px solid var(--dsw-alias-border-l2-darkmode-thin);color:var(--dsw-alias-label-secondary);font:inherit;cursor:pointer;background:0 0;border-radius:999px;align-items:center;gap:6px;padding:7px 13px;font-size:12px;line-height:16px;display:inline-flex}._2S_x-q_modeChip:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}._2S_x-q_modeChip[data-selected=true]{border-color:color-mix(in srgb, var(--dsw-alias-label-primary) 55%, transparent);background:color-mix(in srgb, var(--dsw-alias-label-primary) 12%, transparent);color:var(--dsw-alias-label-primary)}._2S_x-q_templatePanel{box-sizing:border-box;z-index:2;width:100%;max-width:calc(var(--dsh-composer-card-max-width) + 32px);height:var(--office-ppt-template-panel-height,calc(100dvh - 178px));max-height:var(--office-ppt-template-panel-height,calc(100dvh - 178px));background:0 0;flex-direction:column;margin-top:8px;display:flex;position:absolute;top:100%;left:50%;overflow:hidden;transform:translate(-50%)}._2S_x-q_modeRoot[data-placement=fixed]{height:0;overflow:visible}._2S_x-q_modeRoot[data-placement=fixed] ._2S_x-q_templatePanel{max-height:none;top:0}._2S_x-q_modeRoot[data-placement=fixed] ._2S_x-q_templateViewport{overscroll-behavior:contain;flex:auto;padding-bottom:12px;overflow-y:auto}._2S_x-q_templateToolbar{z-index:2;background:0 0;flex:none;justify-content:space-between;align-items:center;gap:12px;min-height:38px;display:flex;position:relative}._2S_x-q_templateViewport{box-sizing:border-box;overscroll-behavior:contain;scrollbar-width:none;min-height:0;padding-bottom:12px;overflow-y:auto}._2S_x-q_templateViewport::-webkit-scrollbar{display:none}._2S_x-q_categoryTabs{scrollbar-width:none;gap:4px;min-width:0;display:flex;overflow-x:auto}._2S_x-q_categoryTabs::-webkit-scrollbar{display:none}._2S_x-q_categoryTabs button{color:var(--dsw-alias-label-caption);font:inherit;cursor:pointer;background:0 0;border:0;border-radius:999px;flex:none;padding:5px 10px;font-size:11px;line-height:16px}._2S_x-q_categoryTabs button:hover{color:var(--dsw-alias-label-primary)}._2S_x-q_categoryTabs button[data-selected=true]{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}._2S_x-q_templateGrid{background:0 0;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px 14px;display:grid}._2S_x-q_templateCard{min-width:0;min-height:0;position:relative}._2S_x-q_templateSelect{width:100%;min-width:0;color:var(--dsw-alias-label-tertiary);font:inherit;cursor:pointer;background:0 0;border:0;outline:0;flex-direction:column;padding:0;display:flex}._2S_x-q_previewViewport,._2S_x-q_selectionPreview{container-type:inline-size}._2S_x-q_previewViewport{aspect-ratio:16/9;border:2px solid #0000;border-radius:9px;width:100%;transition:border-color .12s,transform .12s;display:block;position:relative;overflow:hidden}._2S_x-q_previewViewport:after{content:\"\";position:absolute;inset:0;z-index:2;border-radius:7px;box-shadow:inset 0 0 0 1px color-mix(in srgb, var(--dsw-alias-label-primary) 25%, transparent);pointer-events:none}._2S_x-q_previewViewport>._2S_x-q_preview{aspect-ratio:auto;width:100%;height:100%;animation:.18s _2S_x-q_previewReveal}@keyframes _2S_x-q_previewReveal{0%{opacity:.72}to{opacity:1}}._2S_x-q_templateCard:hover ._2S_x-q_previewViewport{transform:translateY(-1px)}._2S_x-q_templateSelect:active ._2S_x-q_previewViewport{transform:scale(.985)}._2S_x-q_templateCard[data-selected=true] ._2S_x-q_previewViewport{border-color:var(--dsw-alias-label-primary)}._2S_x-q_templateSelect:focus-visible ._2S_x-q_previewViewport{outline:2px solid var(--dsw-alias-label-primary);outline-offset:2px}._2S_x-q_templateName{text-align:center;text-overflow:ellipsis;white-space:nowrap;padding-top:5px;font-size:12px;line-height:16px;overflow:hidden}._2S_x-q_templateCard[data-selected=true] ._2S_x-q_templateName{color:var(--dsw-alias-label-primary)}._2S_x-q_panelState{min-height:160px;color:var(--dsw-alias-label-tertiary);place-items:center;gap:10px;font-size:12px;display:grid}._2S_x-q_templateError{border:1px solid color-mix(in srgb, var(--dsw-alias-state-error-primary) 40%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 10%, transparent);color:var(--dsw-alias-state-error-primary);border-radius:8px;margin-bottom:8px;padding:7px 10px;font-size:11px;line-height:16px}._2S_x-q_retryButton{border:1px solid var(--dsw-alias-border-l2-darkmode-thin);color:var(--dsw-alias-label-secondary);font:inherit;cursor:pointer;background:0 0;border-radius:999px;padding:6px 12px}._2S_x-q_retryButton:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}._2S_x-q_selection{align-items:flex-start;display:flex}._2S_x-q_selectionPreview{aspect-ratio:16/9;transform-origin:50%;flex:none;width:clamp(76px,10vw,92px);transition:transform .2s cubic-bezier(.34,1.56,.64,1);display:block;position:relative;transform:rotate(-3deg)scale(1)}._2S_x-q_selectionFrame{border:2px solid var(--dsw-alias-label-primary);background:var(--dsw-alias-bg-base);border-radius:9px;width:100%;height:100%;display:block;overflow:hidden;box-shadow:0 4px 14px #0000002e}._2S_x-q_selectionRemove{z-index:2;color:#fff;opacity:0;pointer-events:none;cursor:pointer;background:#111111e0;border:1px solid #ffffff4d;border-radius:50%;place-items:center;width:20px;height:20px;padding:0;font:500 15px/1 ui-sans-serif,system-ui,sans-serif;transition:opacity .12s,transform .12s;display:grid;position:absolute;top:-8px;right:-8px;transform:rotate(3deg)scale(.88);box-shadow:0 3px 10px #00000047}._2S_x-q_selectionPreview:hover ._2S_x-q_selectionRemove,._2S_x-q_selectionPreview:focus-within ._2S_x-q_selectionRemove{opacity:1;pointer-events:auto;transform:rotate(3deg)scale(1)}._2S_x-q_selectionRemove:focus-visible{outline:2px solid var(--dsw-alias-label-primary);outline-offset:2px}._2S_x-q_preview{box-sizing:border-box;aspect-ratio:16/9;background:var(--office-bg);color:var(--office-text);flex-direction:column;justify-content:center;padding:12% 10%;display:flex;position:relative;overflow:hidden}._2S_x-q_preview[data-source-preview=true]{background:#fff;padding:0}._2S_x-q_preview[data-source-preview=true]:before,._2S_x-q_preview[data-source-preview=true]:after{display:none}._2S_x-q_previewImage{object-fit:cover;width:100%;height:100%;display:block;position:absolute;inset:0}._2S_x-q_preview:before,._2S_x-q_preview:after{content:\"\";pointer-events:none;position:absolute}._2S_x-q_preview:before{border-left:1px solid color-mix(in srgb, var(--office-accent) 35%, transparent);background:color-mix(in srgb, var(--office-surface) 92%, transparent);width:34%;height:100%;top:0;right:0}._2S_x-q_preview:after{background:var(--office-accent);width:18%;height:clamp(1px,1.3cqi,3px);bottom:13%;right:8%}._2S_x-q_previewAccent{background:var(--office-accent);width:18%;height:clamp(1px,1.3cqi,3px);position:absolute;top:16%;left:10%}._2S_x-q_preview strong,._2S_x-q_preview small{z-index:1;text-overflow:ellipsis;white-space:nowrap;width:66%;display:block;position:relative;overflow:hidden}._2S_x-q_preview strong{margin-top:8%;font-size:clamp(4px,6cqi,14px);font-weight:650;line-height:1.12}._2S_x-q_preview small{color:var(--office-muted);margin-top:4px;font-size:clamp(3px,3.4cqi,8px)}._2S_x-q_preview[data-variant=blue-professional]:before{clip-path:polygon(25% 0,100% 0,100% 100%,0 100%);background:var(--office-surface);border:0;width:40%}._2S_x-q_preview[data-variant=blue-professional]:after{width:15%;height:2px;bottom:14%;right:8%}._2S_x-q_preview[data-variant=editorial-forest]{background:var(--office-surface);color:var(--office-accent);justify-content:flex-end;padding:10%;font-family:Georgia,Times New Roman,serif}._2S_x-q_preview[data-variant=editorial-forest]:before{aspect-ratio:1;border:1px solid var(--office-accent);background:0 0;border-radius:50%;width:16%;height:auto;top:11%;right:9%}._2S_x-q_preview[data-variant=editorial-forest]:after{background:var(--office-accent);width:80%;height:1px;bottom:10%;right:10%}._2S_x-q_preview[data-variant=editorial-forest] ._2S_x-q_previewAccent{background:var(--office-accent);width:26%;height:1px;top:13%}._2S_x-q_preview[data-variant=editorial-forest] strong,._2S_x-q_preview[data-variant=editorial-forest] small{width:78%}._2S_x-q_preview[data-variant=editorial-forest] strong{margin-bottom:11%;font-size:clamp(5px,7.4cqi,17px);font-weight:500}._2S_x-q_preview[data-variant=editorial-forest] small{color:var(--office-bg);font-family:ui-monospace,SFMono-Regular,Menlo,monospace;position:absolute;top:16%;left:10%}._2S_x-q_preview[data-variant=signal]{background:var(--office-bg);color:var(--office-text);font-family:Georgia,Times New Roman,serif}._2S_x-q_preview[data-variant=signal]:before{background-image:linear-gradient(color-mix(in srgb, var(--office-secondary) 55%, transparent) 1px, transparent 1px), linear-gradient(90deg, color-mix(in srgb, var(--office-secondary) 55%, transparent) 1px, transparent 1px);opacity:.45;background-size:18% 31%;border:0;width:100%;height:100%;inset:0}._2S_x-q_preview[data-variant=signal]:after{background:var(--office-secondary);width:80%;height:1px;bottom:15%;left:10%;right:auto}._2S_x-q_preview[data-variant=signal] ._2S_x-q_previewAccent{background:var(--office-accent);width:14%;height:1px;top:18%}._2S_x-q_preview[data-variant=signal] strong{font-size:clamp(5px,6.8cqi,16px);font-weight:600}._2S_x-q_preview[data-variant=signal] small{color:var(--office-muted);letter-spacing:.08em;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}._2S_x-q_preview[data-variant=orange-data]{color:#111;background:#fff;justify-content:flex-start;padding:15% 8%;font-family:Arial,sans-serif}._2S_x-q_preview[data-variant=orange-data]:before{background:linear-gradient(90deg, var(--office-accent) 0 25%, var(--office-secondary) 25% 50%, #f19309 50% 75%, #f9c22b 75%) 0 21% / 100% 28% no-repeat, linear-gradient(#f2f2f2 0 0) 0 70% / 72% 24% no-repeat;border:0;border-top:clamp(1px,.9cqi,2px) solid #111;width:84%;height:45%;inset:8% 8% auto}._2S_x-q_preview[data-variant=orange-data]:after{background:#111;width:84%;height:11%;bottom:11%;right:8%}._2S_x-q_preview[data-variant=orange-data] ._2S_x-q_previewAccent{background:var(--office-accent);width:22%;height:clamp(1px,.9cqi,2px);top:8%;left:8%}._2S_x-q_preview[data-variant=orange-data] strong,._2S_x-q_preview[data-variant=orange-data] small{z-index:2;width:52%}._2S_x-q_preview[data-variant=orange-data] strong{margin-top:10%;font-size:clamp(4px,5.8cqi,13px)}._2S_x-q_preview[data-variant=orange-data] small{color:#5b626d;letter-spacing:.06em;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}._2S_x-q_preview[data-variant=custom]:before{background:linear-gradient(145deg, var(--office-surface), var(--office-secondary));width:40%}._2S_x-q_preview[data-page=\"1\"],._2S_x-q_preview[data-page=\"2\"]{background:color-mix(in srgb, var(--office-bg) 92%, var(--office-surface));color:var(--office-text);font-family:var(--font-sans,ui-sans-serif, system-ui, sans-serif);justify-content:flex-start;padding:8% 8% 7%}._2S_x-q_preview[data-page=\"1\"]:before,._2S_x-q_preview[data-page=\"2\"]:before{background:linear-gradient(90deg, color-mix(in srgb, var(--office-accent) 13%, transparent) 1px, transparent 1px), linear-gradient(color-mix(in srgb, var(--office-accent) 13%, transparent) 1px, transparent 1px);clip-path:none;opacity:.48;background-size:20% 25%;border:0;width:100%;height:100%;inset:0}._2S_x-q_preview[data-page=\"1\"]:after,._2S_x-q_preview[data-page=\"2\"]:after{background:var(--office-accent);width:18%;height:2px;top:8%;bottom:auto;right:8%}._2S_x-q_preview[data-page=\"1\"] ._2S_x-q_previewAccent,._2S_x-q_preview[data-page=\"2\"] ._2S_x-q_previewAccent{background:color-mix(in srgb, var(--office-muted) 38%, transparent);width:84%;height:1px;top:auto;bottom:7%;left:8%}._2S_x-q_preview[data-page=\"1\"] strong,._2S_x-q_preview[data-page=\"2\"] strong{width:78%;margin-top:3%;font-size:clamp(4px,5.3cqi,12px);line-height:1.12}._2S_x-q_previewEyebrow{letter-spacing:.06em;width:74%!important;color:var(--office-muted)!important;margin:0!important;font:500 clamp(3px,2.7cqi,6px)/1.1 ui-monospace,SFMono-Regular,Menlo,monospace!important}._2S_x-q_previewKpis{z-index:1;grid-template-columns:repeat(3,minmax(0,1fr));gap:4px;width:100%;margin-top:auto;display:grid;position:relative}._2S_x-q_previewKpis>span{border:1px solid color-mix(in srgb, var(--office-accent) 18%, transparent);background:color-mix(in srgb, var(--office-surface) 88%, transparent);border-radius:3px;justify-content:space-between;align-items:baseline;min-width:0;padding:7% 8%;display:flex}._2S_x-q_previewKpis b{font-size:clamp(4px,4.8cqi,11px);font-weight:700}._2S_x-q_previewKpis i{color:var(--office-muted);font:clamp(2px,2.2cqi,5px)/1 ui-monospace,monospace}._2S_x-q_previewFooter,._2S_x-q_previewInsight{z-index:1;width:100%;color:var(--office-muted);text-overflow:ellipsis;white-space:nowrap;margin-top:4px;font-size:clamp(2px,2.3cqi,5px);line-height:1.15;display:block;position:relative;overflow:hidden}._2S_x-q_previewChart{z-index:1;border-bottom:1px solid color-mix(in srgb, var(--office-muted) 34%, transparent);align-items:flex-end;gap:7%;height:43%;margin-top:auto;padding:4% 5% 0;display:flex;position:relative}._2S_x-q_previewChart i{background:var(--office-accent);border-radius:2px 2px 0 0;width:11%}._2S_x-q_previewChart i:first-child{opacity:.52;height:28%}._2S_x-q_previewChart i:nth-child(2){opacity:.64;height:46%}._2S_x-q_previewChart i:nth-child(3){opacity:.76;height:41%}._2S_x-q_previewChart i:nth-child(4){opacity:.88;height:72%}._2S_x-q_previewChart i:nth-child(5){height:91%}@media (prefers-reduced-motion:reduce){._2S_x-q_previewViewport,._2S_x-q_selectionPreview{transition:none}._2S_x-q_previewViewport>._2S_x-q_preview{animation:none}}";
		const tagId = "dsh-ppt/OfficePptHero.module.css";
		if (typeof document !== "undefined" && document.querySelector("style[data-plugin-css=" + JSON.stringify(tagId) + "]") === null) {
			const tag = document.createElement("style");
			tag.dataset.plugin = "dsh-ppt";
			tag.dataset.pluginCss = tagId;
			tag.textContent = css + `
.dsh-ppt-panel{order:1;align-self:center;display:flex;flex-direction:column;flex-shrink:0;margin:0;padding:0;box-sizing:border-box;width:calc(100% - 2 * var(--dsh-composer-side-clearance,16px));max-height:400px;overflow:hidden;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#e4e6eb);border-radius:14px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#20232a);box-shadow:0 12px 40px #0000001a,0 2px 8px #00000008;font:13px/1.5 var(--font-sans,system-ui,sans-serif)}
.dsh-ppt-panel-header{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 16px 8px;flex:none}
.dsh-ppt-panel-title{font-size:14px;font-weight:600;margin:0}
.dsh-ppt-panel ._2S_x-q_categoryTabs{padding:0 12px 10px;flex:none;border-bottom:1px solid var(--dsw-alias-border-l2-darkmode-thin,#eceef1)}
.dsh-ppt-panel ._2S_x-q_categoryTabs button{font-size:12px;padding:4px 9px}
.dsh-ppt-panel ._2S_x-q_categoryTabs button[data-selected=true]{color:var(--dsw-alias-label-primary,#20232a);background:var(--dsw-alias-interactive-bg-hover,#f3f4f6)}
.dsh-ppt-panel ._2S_x-q_templateViewport{padding:14px 16px;overflow-y:auto;scrollbar-width:thin;min-height:0;flex:1}
.dsh-ppt-panel ._2S_x-q_templateViewport::-webkit-scrollbar{display:block;width:5px}
.dsh-ppt-panel ._2S_x-q_templateViewport::-webkit-scrollbar-thumb{background:#c5c9d0;border-radius:3px}
.dsh-ppt-panel ._2S_x-q_templateGrid{gap:14px}
.dsh-ppt-panel ._2S_x-q_templateName{font-size:12px;text-align:left;padding:6px 2px 0;width:100%;box-sizing:border-box}
.dsh-ppt-panel ._2S_x-q_previewViewport{border-radius:8px}
.dsh-ppt-selected{position:absolute;top:8px;right:8px;z-index:3;display:grid;place-items:center;width:20px;height:20px;border-radius:50%;background:var(--dsw-alias-label-primary,#20232a);color:var(--dsw-alias-bg-base,#fff);font-size:12px;box-shadow:0 1px 4px #0002;pointer-events:none}
.dsh-ppt-panel button:focus-visible{outline:2px solid var(--dsw-alias-label-primary,#20232a);outline-offset:2px}
.dsh-ppt-panel button:disabled{cursor:wait;opacity:.55}
.dsh-ppt-panel-notice{margin:0 16px 8px;color:var(--dsw-alias-label-secondary,#626875);font-size:12px}
@media(max-width:600px){.dsh-ppt-panel ._2S_x-q_templateGrid{grid-template-columns:repeat(2,minmax(0,1fr))}}
`;
			document.head.appendChild(tag);
		}
		var OfficePptHero_module_css_default = {
			"categoryTabs": "_2S_x-q_categoryTabs",
			"modeChip": "_2S_x-q_modeChip",
			"modeRoot": "_2S_x-q_modeRoot",
			"modeRow": "_2S_x-q_modeRow",
			"panelState": "_2S_x-q_panelState",
			"preview": "_2S_x-q_preview",
			"previewAccent": "_2S_x-q_previewAccent",
			"previewChart": "_2S_x-q_previewChart",
			"previewEyebrow": "_2S_x-q_previewEyebrow",
			"previewFooter": "_2S_x-q_previewFooter",
			"previewImage": "_2S_x-q_previewImage",
			"previewInsight": "_2S_x-q_previewInsight",
			"previewKpis": "_2S_x-q_previewKpis",
			"previewReveal": "_2S_x-q_previewReveal",
			"previewViewport": "_2S_x-q_previewViewport",
			"retryButton": "_2S_x-q_retryButton",
			"selection": "_2S_x-q_selection",
			"selectionFrame": "_2S_x-q_selectionFrame",
			"selectionPreview": "_2S_x-q_selectionPreview",
			"selectionRemove": "_2S_x-q_selectionRemove",
			"templateCard": "_2S_x-q_templateCard",
			"templateError": "_2S_x-q_templateError",
			"templateGrid": "_2S_x-q_templateGrid",
			"templateName": "_2S_x-q_templateName",
			"templatePanel": "_2S_x-q_templatePanel",
			"templateSelect": "_2S_x-q_templateSelect",
			"templateToolbar": "_2S_x-q_templateToolbar",
			"templateViewport": "_2S_x-q_templateViewport"
		};
		//#endregion
		//#region ../office-ppt/src/client/OfficePptHero.tsx
		/** DSH PPT template chooser integrated into the blank-session composer. */
		const TEMPLATE_CATEGORIES = [
			"all",
			"strategy",
			"business",
			"consulting",
			"finance",
			"work",
			"promotion",
			"academic", "editorial"
		];
		const TEMPLATE_LOAD_TIMEOUT_MS = 8e3;
		const FALLBACK_TEMPLATE_PREVIEW_PAGE_COUNT = 3;
		const TEMPLATE_PREVIEW_AUTOPLAY_START_MS = 420;
		const TEMPLATE_PREVIEW_AUTOPLAY_INTERVAL_MS = 900;
		const TEMPLATE_PREVIEW_WHEEL_LOCK_MS = 180;
		const EMPTY_STATE = {
			activeMode: null,
			pickerOpen: false,
			loading: false,
			templates: [],
			selectedId: null,
            notice: false,
			error: ""
		};
		/** Session-keyed presentation state shared by the toolbar button and template panel. */
		var OfficePptHeroStore = class {
			constructor() {
				this.states = /* @__PURE__ */ new Map();
				this.listeners = /* @__PURE__ */ new Map();
			}
			snapshot(sessionId) {
				return this.states.get(sessionId) ?? EMPTY_STATE;
			}
			subscribe(sessionId, listener) {
				const listeners = this.listeners.get(sessionId) ?? /* @__PURE__ */ new Set();
				listeners.add(listener);
				this.listeners.set(sessionId, listeners);
				return () => {
					listeners.delete(listener);
					if (listeners.size === 0) this.listeners.delete(sessionId);
				};
			}
			setPickerOpen(sessionId, pickerOpen) {
				this.update(sessionId, current => ({ ...current, pickerOpen }));
			}
			setMode(sessionId, activeMode) {
				this.update(sessionId, (current) => ({
					...current,
					activeMode,
					error: ""
				}));
			}
			setLoading(sessionId, loading) {
				this.update(sessionId, (current) => ({
					...current,
					loading,
					error: loading ? "" : current.error
				}));
			}
			setNotice(sessionId, notice) {
                this.update(sessionId, current => ({ ...current, notice }));
            }
            setError(sessionId, error) {
				this.update(sessionId, (current) => ({
					...current,
					loading: false,
					error
				}));
			}
			setTemplates(sessionId, templates) {
				this.update(sessionId, (current) => ({
					...current,
					loading: false,
					templates,
					selectedId: templates.some((template) => template.id === current.selectedId) ? current.selectedId : null,
					error: ""
				}));
			}
			select(sessionId, template, activeMode) {
				this.update(sessionId, (current) => ({
					...current,
					activeMode,
					loading: false,
					selectedId: template.id,
                    notice: false,
					error: ""
				}));
			}
			deselect(sessionId, activeMode) {
				this.update(sessionId, (current) => ({
					...current,
					activeMode,
					loading: false,
					selectedId: null,
					error: ""
				}));
			}
			update(sessionId, transform) {
				this.states.set(sessionId, transform(this.snapshot(sessionId)));
				for (const listener of this.listeners.get(sessionId) ?? []) listener();
			}
		};
		function useMode(mode, sessionId) {
			return (0, react.useSyncExternalStore)((listener) => mode.subscribe(sessionId, listener), () => mode.snapshot(sessionId), () => mode.snapshot(sessionId));
		}
		function templateStyle(template) {
			return {
				"--office-bg": `#${template.palette.background}`,
				"--office-surface": `#${template.palette.surface}`,
				"--office-text": `#${template.palette.text}`,
				"--office-muted": `#${template.palette.muted}`,
				"--office-accent": `#${template.palette.accent}`,
				"--office-secondary": `#${template.palette.secondary}`
			};
		}
		function templateCategory(template) {
			if (template.origin === "extracted") return "custom";
			return template.category ?? "business";
		}
		function templateVariant(template) {
			return template.source?.visualGrammar ?? "custom";
		}
		function templatePreviewPages(template) {
			return CURATED_TEMPLATE_PREVIEWS[template.id]?.length ?? FALLBACK_TEMPLATE_PREVIEW_PAGE_COUNT;
		}
		function TemplatePreview({ template, page = 0 }) {
			const previewImage = CURATED_TEMPLATE_PREVIEWS[template.id]?.[page];
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
				className: OfficePptHero_module_css_default.preview,
				"data-page": page,
				"data-source-preview": previewImage === void 0 ? void 0 : true,
				"data-variant": templateVariant(template),
				style: templateStyle(template),
				children: [previewImage !== void 0 && /* @__PURE__ */ (0, react_jsx_runtime.jsx)("img", {
					className: OfficePptHero_module_css_default.previewImage,
					src: previewImage,
					loading: "lazy",
					decoding: "async",
					alt: "",
					"aria-hidden": "true"
				}), previewImage === void 0 && /* @__PURE__ */ (0, react_jsx_runtime.jsxs)(react_jsx_runtime.Fragment, { children: [
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { className: OfficePptHero_module_css_default.previewAccent }),
					page === 0 && /* @__PURE__ */ (0, react_jsx_runtime.jsxs)(react_jsx_runtime.Fragment, { children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: template.previewTitle }), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", { children: template.previewSubtitle })] }),
					page === 1 && /* @__PURE__ */ (0, react_jsx_runtime.jsxs)(react_jsx_runtime.Fragment, { children: [
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", {
							className: OfficePptHero_module_css_default.previewEyebrow,
							children: template.previewSubtitle
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: template.previewTitle }),
						/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
							className: OfficePptHero_module_css_default.previewKpis,
							"aria-hidden": "true",
							children: [
								/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", { children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("b", { children: "32%" }), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", { children: "01" })] }),
								/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", { children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("b", { children: "18.6" }), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", { children: "02" })] }),
								/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", { children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("b", { children: "04" }), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", { children: "03" })] })
							]
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
							className: OfficePptHero_module_css_default.previewFooter,
							children: template.description
						})
					] }),
					page === 2 && /* @__PURE__ */ (0, react_jsx_runtime.jsxs)(react_jsx_runtime.Fragment, { children: [
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", {
							className: OfficePptHero_module_css_default.previewEyebrow,
							children: template.previewSubtitle
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: template.previewTitle }),
						/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
							className: OfficePptHero_module_css_default.previewChart,
							"aria-hidden": "true",
							children: [
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", {}),
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", {}),
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", {}),
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", {}),
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("i", {})
							]
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
							className: OfficePptHero_module_css_default.previewInsight,
							children: template.description
						})
					] })
				] })]
			});
		}
		function TemplatePreviewDeck({ template, page }) {
			return /* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
				className: OfficePptHero_module_css_default.previewViewport,
				children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)(TemplatePreview, {
					template,
					page
				}, page)
			});
		}
		function TemplateCard({ template, selected, choose, disabled }) {
			const [page, setPage] = (0, react.useState)(0);
			const [hovered, setHovered] = (0, react.useState)(false);
			const card = (0, react.useRef)(null);
			const hoveredCard = (0, react.useRef)(false);
			const lastWheelAt = (0, react.useRef)(Number.NEGATIVE_INFINITY);
			const pageCount = templatePreviewPages(template);
			(0, react.useEffect)(() => {
				if (!hovered || pageCount <= 1 || reducedMotionPreferred()) return;
				const advance = () => {
					if (!hoveredCard.current || document.visibilityState === "hidden") return;
					if (performance.now() - lastWheelAt.current < TEMPLATE_PREVIEW_AUTOPLAY_INTERVAL_MS) return;
					setPage((current) => (current + 1) % pageCount);
				};
				let interval;
				const start = window.setTimeout(() => {
					advance();
					interval = window.setInterval(advance, TEMPLATE_PREVIEW_AUTOPLAY_INTERVAL_MS);
				}, TEMPLATE_PREVIEW_AUTOPLAY_START_MS);
				return () => {
					window.clearTimeout(start);
					if (interval !== void 0) window.clearInterval(interval);
				};
			}, [hovered, pageCount]);
			(0, react.useEffect)(() => {
				const element = card.current;
				if (!hovered || element === null) return;
				const stopOutsideCard = (event) => {
					if (event.target instanceof Node && element.contains(event.target)) return;
					hoveredCard.current = false;
					setHovered(false);
					setPage(0);
					lastWheelAt.current = Number.NEGATIVE_INFINITY;
				};
				document.addEventListener("mousemove", stopOutsideCard, true);
				return () => {
					document.removeEventListener("mousemove", stopOutsideCard, true);
				};
			}, [hovered]);
			(0, react.useEffect)(() => {
				const element = card.current;
				if (element === null || pageCount <= 1) return;
				const pageWithWheel = (event) => {
					const horizontalDistance = Math.abs(event.deltaX);
					const verticalDistance = Math.abs(event.deltaY);
					if (horizontalDistance < 4 || horizontalDistance < verticalDistance * .65) return;
					const delta = event.deltaX;
					if (!Number.isFinite(delta) || delta === 0) return;
					event.preventDefault();
					const now = performance.now();
					if (now - lastWheelAt.current < TEMPLATE_PREVIEW_WHEEL_LOCK_MS) return;
					lastWheelAt.current = now;
					setPage((current) => (current + (delta > 0 ? 1 : -1) + pageCount) % pageCount);
				};
				element.addEventListener("wheel", pageWithWheel, { passive: false });
				return () => {
					element.removeEventListener("wheel", pageWithWheel);
				};
			}, [pageCount]);
			return /* @__PURE__ */ (0, react_jsx_runtime.jsx)("article", {
				ref: card,
				className: OfficePptHero_module_css_default.templateCard,
				"data-selected": selected || void 0,
				onMouseEnter: () => {
					hoveredCard.current = true;
					setHovered(true);
				},
				onMouseLeave: () => {
					hoveredCard.current = false;
					setHovered(false);
					setPage(0);
					lastWheelAt.current = Number.NEGATIVE_INFINITY;
				},
				children: /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("button", {
					type: "button",
					className: OfficePptHero_module_css_default.templateSelect,
					"aria-label": template.name,
					"aria-pressed": selected,
					disabled,
					onClick: () => {
						choose(template);
					},
					children: [
						selected && react_jsx_runtime.jsx("span", { className: "dsh-ppt-selected", "aria-hidden": true, children: "✓" }),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)(TemplatePreviewDeck, {
							template,
							page
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
							className: OfficePptHero_module_css_default.templateName,
							children: template.name
						})
					]
				})
			});
		}
		function reducedMotionPreferred() {
			return typeof window.matchMedia === "function" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
		}
		function loadTemplateState(client, timeoutMessage) {
			const controller = new AbortController();
			return new Promise((resolve, reject) => {
				const timer = window.setTimeout(() => {
					controller.abort();
					reject(new Error(timeoutMessage));
				}, TEMPLATE_LOAD_TIMEOUT_MS);
				client.call("state", {}, controller.signal).then((next) => {
					window.clearTimeout(timer);
					resolve(next);
				}, (reason) => {
					window.clearTimeout(timer);
					reject(reason instanceof Error ? reason : new Error(String(reason)));
				});
			});
		}
		/** Monochrome presentation screen used by both PPT mode controls. */
		function DesktopPptIcon() {
			return (0, react_jsx_runtime.jsxs)("svg", {
				className: "desktop-ppt-icon", width: 18, height: 18, viewBox: "0 0 24 24",
				fill: "none", stroke: "currentColor", strokeWidth: 1.75,
				strokeLinecap: "round", strokeLinejoin: "round", "aria-hidden": true, focusable: false,
				children: [
					(0, react_jsx_runtime.jsx)("path", { d: "M2 3h20M21 3v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3M7 21l5-4 5 4M12 17v4" })
				]
			});
		}
		/** Template picker shown only for the native PPT agent preset. */
		function OfficePptChooser({ client, mode, sessionId, t }) {
			const state = useMode(mode, sessionId);
			const [category, setCategory] = react.useState("all");

			react.useEffect(() => {
				if (mode.snapshot(sessionId).loading || state.templates.length > 0 || state.error !== "") return;
				mode.setLoading(sessionId, true);
				loadTemplateState(client, t("templates.loadTimeout")).then(next => {
					mode.setTemplates(sessionId, next.templates);
					const template = next.templates.find(item => item.id === next.selectedTemplateId);
					if (template) mode.select(sessionId, template, "ppt");
					else mode.deselect(sessionId, "ppt");
					mode.setNotice(sessionId, next.templateMigration?.reason === "template-retired");
				}).catch(reason => mode.setError(sessionId, reason instanceof Error ? reason.message : String(reason)));
			}, [client, mode, sessionId, state.loading, state.templates.length, state.error, t]);

			const choose = async template => {
				if (state.loading || template.id === state.selectedId) return;
				mode.setLoading(sessionId, true);
				try {
					await client.call("template/select", { templateId: template.id, mode: "ppt" });
					mode.select(sessionId, template, "ppt");
				} catch (reason) {
					mode.setError(sessionId, reason instanceof Error ? reason.message : String(reason));
				}
			};
			const templates = state.templates.filter(template => templateSupportsMode(template, "ppt"));
			const categories = ["all", ...new Set(templates.map(templateCategory))];
			const visible = templates.filter(template => category === "all" || templateCategory(template) === category);
			const h = react.createElement;
			return h("section", {
					"aria-label": t("templates.title"),
					className: "dsh-ppt-panel", "data-office-ppt-template-panel": "", "data-placement": "bottom"
				},
					h("div", { className: "dsh-ppt-panel-header" },
						h("h2", { className: "dsh-ppt-panel-title" }, t("templates.title"))),
					state.notice && h("p", { className: "dsh-ppt-panel-notice", role: "status" }, t("templates.migrated")),
					h("div", { className: OfficePptHero_module_css_default.categoryTabs, role: "group", "aria-label": t("templates.categories") },
						categories.map(item => h("button", {
							key: item, type: "button", "data-ppt-category": item, "aria-pressed": category === item,
							"data-selected": category === item || undefined, onClick: () => setCategory(item)
						}, t(`templates.category.${item}`)))),
					h("div", { className: OfficePptHero_module_css_default.templateViewport, "data-office-ppt-template-viewport": "", "data-native-wheel-owner": "", "aria-busy": state.loading },
						state.error && h("div", { className: OfficePptHero_module_css_default.templateError, role: "alert" }, state.error,
							state.templates.length === 0 && h("button", { type: "button", className: OfficePptHero_module_css_default.retryButton, onClick: () => mode.setError(sessionId, "") }, t("templates.retry"))),
						state.loading && templates.length === 0 ? h("div", { className: OfficePptHero_module_css_default.panelState }, t("status.loading")) :
						h("div", { className: OfficePptHero_module_css_default.templateGrid }, visible.map(template => h(TemplateCard, {
							key: template.id, template, selected: template.id === state.selectedId, choose, disabled: state.loading
						})))));
		}
		function OfficePptComposer(props) {
			const blank = props.useSession(current => current.blank);
			const agentPreset = props.useSessions(current => {
				const summary = current.byId[props.sessionId];
				return summary?.projectionValues?.agentPreset ?? summary?.agentPreset;
			});
			if (!blank || agentPreset !== "ppt") return null;
			return react_jsx_runtime.jsx(OfficePptChooser, {
				client: props.client, mode: props.mode, sessionId: props.sessionId, t: props.t
			}, props.sessionId);
		}
		//#endregion
		//#region ../office-ppt/src/client/locales.ts
		/** Office PPT composer-mode dictionaries. */
		/** Locale namespace. */
		const NS = "dsh-ppt";
		/** Simplified Chinese dictionary. */
		const zh = {
			"mode.label": "PPT",
			"mode.region": "演示文稿模式",
			"mode.exit": "退出 PPT 模式",
			"templates.close": "收起模板选择",
			"composer.selectedTemplate": "已选模板",
			"composer.removeTemplate": "取消选择模板",
			"templates.title": "选择模板",
            "templates.migrated": "原模板已下架，已切换为工程蓝图。已有文件不受影响。",
			"templates.categories": "模板分类",
			"templates.empty": "该分类下还没有模板",
			"templates.loadTimeout": "模板加载超时，请重试",
			"templates.retry": "重新加载模板",
			"templates.category.all": "全部",
			"templates.category.custom": "自定义",
			"templates.category.business": "商务",
			"templates.category.strategy": "策略",
			"templates.category.consulting": "咨询",
			"templates.category.finance": "金融",
			"templates.category.work": "工作",
			"templates.category.promotion": "推广",
			"templates.category.academic": "学术",
			"templates.category.data": "数据",
			"templates.category.editorial": "编辑",
			"templates.category.briefing": "简报",
			"status.loading": "正在处理…",
			"tool.structure": "PPT 结构",
			"tool.structure.open": "查看第 {page} 页：{title}",
			"tool.pages": "页",
			"tool.inspect": "详情",
			"tool.plan": "PPT 规划与校验",
			"tool.plan.story": "叙事",
			"tool.plan.layouts": "种版式",
			"tool.plan.qa": "页面",
			"tool.preview": "PPT 预览",
			"tool.preview.loading": "正在读取最终版本…",
			"tool.preview.failed": "最终版本读取失败",
			"tool.preview.previous": "上一页",
			"tool.preview.next": "下一页",
			"tool.preview.slide": "第 {page} 页，共 {total} 页",
			"tool.preview.retry": "重试",
			"tool.preview.revision": "只读 · 修订版 {revision}",
			"tool.preview.rendering": "正在渲染最终 PPTX…",
			"tool.preview.structure": "内容结构预览",
			"tool.preview.structureOnly": "当前显示内容结构，打开 PPT 可查看最终样式。",
			"tool.preview.structureFailed": "最终 PPTX 渲染失败，当前显示内容结构。",
			"tool.preview.structureLarge": "文件超过在线预览大小上限，当前显示内容结构。",
			"tool.artifact": "PPT 产出物",
			"tool.artifact.saved": "已保存到当前工作区",
			"tool.artifact.fallback": "可编辑演示文稿.pptx",
			"tool.download": "下载 PPTX",
			"tool.openFile": "打开 PPT",
			"tool.openFolder": "在 Finder 中显示",
			"tool.downloading": "正在下载…",
			"tool.openingFile": "正在打开…",
			"tool.openingFolder": "正在打开…",
			"tool.fileActionFailed": "文件操作失败，请重试",
			"tool.create.running": "正在生成 PPT",
			"tool.create.done": "PPT 已生成",
			"tool.create.failed": "PPT 生成失败",
			"tool.update.running": "正在修改第 {page} 页",
			"tool.update.done": "第 {page} 页已更新",
			"tool.update.failed": "第 {page} 页修改失败",
			"tool.update.content": "页面内容已更新"
		};
		/** English dictionary. */
		const en = {
			"mode.label": "PPT",
			"mode.region": "Presentation modes",
			"mode.exit": "Exit PPT mode",
			"templates.close": "Close template picker",
			"composer.selectedTemplate": "Selected template",
			"composer.removeTemplate": "Remove selected template",
			"templates.title": "Select template",
            "templates.migrated": "The previous template was retired. Engineering Blueprint is selected; existing files are unchanged.",
			"templates.categories": "Template categories",
			"templates.empty": "No templates in this category",
			"templates.loadTimeout": "Template loading timed out. Try again.",
			"templates.retry": "Reload templates",
			"templates.category.all": "All",
			"templates.category.custom": "Custom",
			"templates.category.business": "Business",
			"templates.category.strategy": "Strategy",
			"templates.category.consulting": "Consulting",
			"templates.category.finance": "Finance",
			"templates.category.work": "Work",
			"templates.category.promotion": "Promotion",
			"templates.category.academic": "Academic",
			"templates.category.data": "Data",
			"templates.category.editorial": "Editorial",
			"templates.category.briefing": "Briefing",
			"status.loading": "Processing…",
			"tool.structure": "Presentation structure",
			"tool.structure.open": "View slide {page}: {title}",
			"tool.pages": "slides",
			"tool.inspect": "Details",
			"tool.plan": "Presentation planning and QA",
			"tool.plan.story": "Story",
			"tool.plan.layouts": "layouts",
			"tool.plan.qa": "Pages",
			"tool.preview": "Presentation preview",
			"tool.preview.loading": "Loading the final revision…",
			"tool.preview.failed": "Could not load the final revision",
			"tool.preview.previous": "Previous slide",
			"tool.preview.next": "Next slide",
			"tool.preview.slide": "Slide {page} of {total}",
			"tool.preview.retry": "Retry",
			"tool.preview.revision": "Read only · Revision {revision}",
			"tool.preview.rendering": "Rendering the final PPTX…",
			"tool.preview.structure": "Content structure preview",
			"tool.preview.structureOnly": "This view shows content structure. Open the PPTX for final styling.",
			"tool.preview.structureFailed": "Final PPTX rendering failed. This view shows content structure.",
			"tool.preview.structureLarge": "This file exceeds the inline preview limit. This view shows content structure.",
			"tool.artifact": "Presentation output",
			"tool.artifact.saved": "Saved in the current workspace",
			"tool.artifact.fallback": "Editable presentation.pptx",
			"tool.download": "Download PPTX",
			"tool.openFile": "Open presentation",
			"tool.openFolder": "Show in Finder",
			"tool.downloading": "Downloading…",
			"tool.openingFile": "Opening…",
			"tool.openingFolder": "Opening…",
			"tool.fileActionFailed": "File action failed. Try again.",
			"tool.create.running": "Generating presentation",
			"tool.create.done": "Presentation generated",
			"tool.create.failed": "Presentation failed",
			"tool.update.running": "Updating slide {page}",
			"tool.update.done": "Slide {page} updated",
			"tool.update.failed": "Slide {page} update failed",
			"tool.update.content": "Slide content updated"
		};
		//#endregion
		//#region ../office-ppt/src/client/rpc.ts
		/**
		* Bind every request to the active conversation session.
		* @param rpc - Browser Connection RPC client.
		* @param sessionId - Active conversation session.
		* @returns Session-bound Office PPT client.
		*/
		function createOfficePptClient(rpc, sessionId) {
			return { async call(endpoint, payload = {}, signal) {
				const outer = await rpc.call("/dsh-ppt", endpoint, {
					sessionId,
					...payload
				}, signal);
				if (!outer.ok) throw new Error(outer.error.message);
				const raw = outer.value;
				if (raw === null || typeof raw !== "object" || !("status" in raw)) throw new Error("Office PPT returned an invalid response");
				const inner = raw;
				if (inner.status === "error") throw new Error(inner.error.message);
				return inner.data;
			} };
		}
		//#endregion
		//#region ../office-ppt/src/client/standard.ts
		/** Browser services required by the standard Composer adapter. */
		const standardInject = [
			"slots",
			"locale",
			"connection"
		];
		/** Build the session-scoped browser face shared by both Composer layouts. */
		function officePptHeroInjection(ctx) {
			ctx.effect(() => ctx.locale.register(NS, {
				zh,
				en
			}), "dsh-ppt: dictionaries");
			const connection = ctx.get("connection");
			const mode = new OfficePptHeroStore();
			return (sessionId) => ({
				client: createOfficePptClient(connection.rpc, sessionId),
				mode
			});
		}
		/** Register the template picker below the composer for the native PPT preset. */
		function applyStandard(ctx) {
			const injectHero = officePptHeroInjection(ctx);
			const name = "conversation.input.dock";
			ctx.slots.inject(name, () => ctx.slots.register({
				name, id: "dsh-ppt", order: 20, locale: NS, inject: injectHero
			}, props => react_jsx_runtime.jsx(OfficePptComposer, {
				client: props.client, mode: props.mode, sessionId: props.sessionId,
				useSession: props.useSession, useSessions: props.useSessions, t: props.t
			}, props.sessionId)));
		}
		//#endregion
		//#region src/client/index.ts
		/** Required standard browser services. */
		const inject = standardInject;
		/** Mount the native-PPT-preset template picker below the input card. */
		function apply(ctx) {
			applyStandard(ctx);
		}
		//#endregion
		exports.apply = apply;
		exports.inject = inject;
		return module.exports;
	}
});
