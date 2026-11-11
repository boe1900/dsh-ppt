/** Fit the entire slide into the measured viewport; zoom multiplies that fit. */
function FitPreviewImage({ src, alt, zoom, onZoom, onError, className }) {
  const h = react.createElement;
  const viewport = react.useRef(null);
  const imageRef = react.useRef(null);
  const [size, setSize] = react.useState({ width: 1, height: 1 });
  const [natural, setNatural] = react.useState({ width: 16, height: 9 });
  const imageReady = image => {
    if (!image?.naturalWidth || !image.naturalHeight) return;
    setNatural({ width: image.naturalWidth, height: image.naturalHeight });
  };
  // Cached images may finish before React installs the load handler.
  react.useEffect(() => { if (imageRef.current?.complete) imageReady(imageRef.current); }, [src]);
  react.useLayoutEffect(() => {
    const node = viewport.current;
    const measure = () => setSize({ width: Math.max(1, node.clientWidth), height: Math.max(1, node.clientHeight) });
    measure();
    if (typeof ResizeObserver === 'undefined') { window.addEventListener('resize', measure); return () => window.removeEventListener('resize', measure); }
    const observer = new ResizeObserver(measure); observer.observe(node); return () => observer.disconnect();
  }, []);
  const zoomValue = react.useRef(zoom);
  zoomValue.current = zoom;
  react.useEffect(() => {
    const node=viewport.current;
    const clamp=value=>Math.max(1,Math.min(3,value));
    const wheel=event=>{
      if(!event.ctrlKey) return;
      event.preventDefault();
      onZoom?.(clamp(zoomValue.current*Math.exp(-event.deltaY*.01)));
    };
    let pinch;
    const distance=touches=>Math.hypot(touches[0].clientX-touches[1].clientX,touches[0].clientY-touches[1].clientY);
    const start=event=>{if(event.touches.length===2){event.preventDefault();pinch={distance:distance(event.touches),zoom:zoomValue.current};}};
    const move=event=>{if(pinch && event.touches.length===2){event.preventDefault();if(pinch.distance>0) onZoom?.(clamp(pinch.zoom*distance(event.touches)/pinch.distance));}};
    const end=()=>{pinch=undefined;};
    node.addEventListener('wheel',wheel,{passive:false});
    node.addEventListener('touchstart',start,{passive:false});node.addEventListener('touchmove',move,{passive:false});
    node.addEventListener('touchend',end);node.addEventListener('touchcancel',end);
    return ()=>{node.removeEventListener('wheel',wheel);node.removeEventListener('touchstart',start);node.removeEventListener('touchmove',move);node.removeEventListener('touchend',end);node.removeEventListener('touchcancel',end);};
  },[onZoom]);
  const scale = Math.min(size.width / natural.width, size.height / natural.height) * zoom;
  const width = natural.width * scale, height = natural.height * scale;
  return h('div', { ref: viewport, className: `ppt-fit-viewport ${className ?? ''}`, style: { flex: 1, minHeight: 0, height: '100%', overflow: 'auto' } },
    h('div', { style: { width: Math.max(size.width, width), height: Math.max(size.height, height), display: 'grid', placeItems: 'center' } },
      h('img', { ref: imageRef, src, alt, 'data-zoom': zoom, onError, onLoad: event => imageReady(event.currentTarget),
        style: { display: 'block', width, height, maxWidth: 'none', flex: 'none' } })));
}

function personalPreviewSlides(template) {
  return template.templateSamples?.representativePages?.length
    ? template.templateSamples.representativePages : Array.from({length:template.slideCount},(_,i)=>i+1);
}

function TemplatePreviewViewer({ title, page, count, navigate, close, getThumbnail, loadThumbnail, t, children, action, loading }) {
  const h = react.createElement;
  const dialog = react.useRef(null);
  const strip = react.useRef(null);
  const titleId = react.useId();
  const [thumbnails, setThumbnails] = react.useState({});
  const sources = react.useRef({ getThumbnail, loadThumbnail });
  sources.current = { getThumbnail, loadThumbnail };
  const currentThumbnail = getThumbnail(page);
  react.useEffect(() => {
    if (currentThumbnail) setThumbnails(value => ({ ...value, [page]: { src: currentThumbnail } }));
  }, [page, currentThumbnail]);
  react.useLayoutEffect(() => {
    const node = dialog.current;
    node.showModal(); node.querySelector('button')?.focus();
    return () => { if (node.open) node.close(); };
  }, []);
  react.useEffect(() => {
    strip.current?.querySelector('[aria-current="page"]')?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }, [page]);
  react.useEffect(() => {
    if (!loadThumbnail || !window.IntersectionObserver) return;
    let alive = true, active = 0;
    const queue = [], requested = new Set();
    const pump = () => {
      while (alive && active < 2 && queue.length) {
        const index = queue.shift();
        if (sources.current.getThumbnail(index)) continue;
        active++;
        Promise.resolve().then(() => sources.current.loadThumbnail(index))
          .then(src => { if (alive) setThumbnails(value => ({ ...value, [index]: { src } })); })
          .catch(() => { if (alive) setThumbnails(value => ({ ...value, [index]: { failed: true } })); })
          .finally(() => { active--; pump(); });
      }
    };
    const observer = new window.IntersectionObserver(entries => {
      for (const entry of entries) {
        const index = Number(entry.target.dataset.page);
        if (!entry.isIntersecting || requested.has(index)) continue;
        requested.add(index); queue.push(index); observer.unobserve(entry.target);
      }
      pump();
    }, { root: strip.current, rootMargin: '0px 100px' });
    strip.current.querySelectorAll('[data-page]').forEach(node => observer.observe(node));
    return () => { alive = false; observer.disconnect(); };
  }, [count, Boolean(loadThumbnail)]);
  return h('dialog', { ref: dialog, className: 'ppt-template-viewer', 'aria-labelledby': titleId,
    onCancel: event => { event.preventDefault(); close(); },
    onClick: event => {
      if (event.target !== event.currentTarget) return;
      const rect = event.currentTarget.getBoundingClientRect();
      if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) close();
    },
    onKeyDown: event => {
      if (event.defaultPrevented) return;
      if (!event.target.closest('.ppt-template-thumbnails')) {
        if (event.key === 'ArrowLeft' && page > 0) { event.preventDefault(); navigate(page - 1); }
        if (event.key === 'ArrowRight' && page < count - 1) { event.preventDefault(); navigate(page + 1); }
      }
      if (event.key === 'Tab') {
        const fields = [...event.currentTarget.querySelectorAll('button:not(:disabled):not([tabindex="-1"])')];
        if (event.shiftKey && document.activeElement === fields[0]) { event.preventDefault(); fields.at(-1)?.focus(); }
        else if (!event.shiftKey && document.activeElement === fields.at(-1)) { event.preventDefault(); fields[0]?.focus(); }
      }
    } },
    h('header', { className: 'ppt-template-viewer-header' },
      h('h2', { id: titleId, title }, title),
      h('span', { className: 'ppt-template-page', role: 'status', 'aria-live': 'polite' }, `${count ? page + 1 : 0} / ${count}`),
      h('button', { type: 'button', className: 'ppt-template-close', 'aria-label': t('personal.close'), onClick: close }, '×')),
    h('div', { className: 'ppt-template-canvas' },
      h('button', { type: 'button', className: 'ppt-template-arrow ppt-template-previous', disabled: page <= 0, 'aria-label': t('personal.previous'), onClick: () => navigate(page - 1) }, '‹'),
      h('div', { className: 'ppt-template-stage', 'aria-busy': loading }, children),
      h('button', { type: 'button', className: 'ppt-template-arrow ppt-template-next', disabled: page >= count - 1, 'aria-label': t('personal.next'), onClick: () => navigate(page + 1) }, '›')),
    h('footer', { className: 'ppt-template-viewer-footer' },
      h('div', { ref: strip, className: 'ppt-template-thumbnails', role: 'group', 'aria-label': t('personal.layouts'),
        onKeyDown: event => {
          const current = Number(event.target.closest('[data-page]')?.dataset.page);
          let next;
          if (event.key === 'ArrowLeft') next = Math.max(0, current - 1);
          if (event.key === 'ArrowRight') next = Math.min(count - 1, current + 1);
          if (event.key === 'Home') next = 0;
          if (event.key === 'End') next = count - 1;
          if (next === undefined || !Number.isInteger(current)) return;
          event.preventDefault(); event.stopPropagation(); navigate(next);
          strip.current.querySelector(`[data-page="${next}"]`)?.focus();
        } },
        ...Array.from({ length: count }, (_, index) => {
          const src = thumbnails[index]?.failed ? null : getThumbnail(index) ?? thumbnails[index]?.src;
          return h('button', { key: index, type: 'button', className: 'ppt-template-thumb', 'data-page': index,
            'aria-label': `${t('personal.preview')} ${index + 1}`, 'aria-current': index === page ? 'page' : undefined,
            tabIndex: index === page ? 0 : -1, onClick: () => navigate(index) },
            src && h('img', { src, alt: '', loading: 'lazy', onError: () => setThumbnails(value => ({ ...value, [index]: { failed: true } })) }),
            h('span', null, index + 1));
        })),
      action));
}

/** Shared detailed template viewer; selection remains an independent action. */
function TemplatePreviewCard({ template, selected, choose, client, t, disabled }) {
  const h = react.createElement;
  const [open, setOpen] = react.useState(false);
  const [page, setPage] = react.useState(0);
  const [pages, setPages] = react.useState({});
  const [loading, setLoading] = react.useState(false);
  const [error, setError] = react.useState('');
  const [zoom, setZoom] = react.useState(1);
  const request = react.useRef(0);
  const staticPages = FULL_TEMPLATE_PREVIEWS[template.id] ?? template.previewImages ?? [];
  const detailedPersonal = template.origin === 'personal' && client.bound === true;
  const slides=personalPreviewSlides(template);
  const count = detailedPersonal ? slides.length : staticPages.length;
  const cacheKey=slide=>`${template.version ?? template.previewVersion ?? 0}:${slide}`;
  const src = detailedPersonal ? pages[cacheKey(slides[page])] : staticPages[page];
  async function load(next) {
    const slide=slides[next],key=cacheKey(slide);
    const token = ++request.current;
    setPage(next); setZoom(1); setError('');
    if (!detailedPersonal || pages[key]) { setLoading(false); return; }
    setLoading(true);
    try {
      const result = await client.call('template/preview-saved-page', { templateId: template.id, page: slide });
      if (request.current === token) setPages(value => ({ ...value, [key]: result.preview }));
    } catch (reason) {
      if (request.current === token) setError(reason instanceof Error ? reason.message : String(reason));
    } finally { if (request.current === token) setLoading(false); }
  }
  react.useEffect(() => {
    if (!open) return;
    load(0);
    return () => { request.current++; };
  }, [open, template.id, template.version, template.previewVersion]);
  react.useEffect(() => () => { request.current++; }, [client, template.id, template.version, template.previewVersion]);
  return h('div', { className: 'ppt-preview-card' },
    h('style', null, TEMPLATE_VIEWER_CSS),
    h(TemplateCard, { template, selected, choose, disabled }),
    h('button', { type: 'button', className: 'ppt-preview-open', 'aria-label': `${t('personal.view')} · ${template.name}`, title: t('personal.view'), onClick: () => setOpen(true) },
      h('svg', { width: 16, height: 16, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.6, 'aria-hidden': true },
        h('path', { d: 'M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z' }),
        h('circle', { cx: 12, cy: 12, r: 3 }))),
    open && h(TemplatePreviewViewer, { key: `${template.id}:${template.version ?? template.previewVersion ?? 0}`, title: template.name, page, count, navigate: load, close: () => setOpen(false), t, loading,
      getThumbnail: index => detailedPersonal ? pages[cacheKey(slides[index])] : staticPages[index],
      loadThumbnail: detailedPersonal ? async index => (await client.call('template/preview-saved-page', { templateId: template.id, page: slides[index] })).preview : undefined,
      action: h('button', { type: 'button', className: 'ppt-preview-use', disabled, onClick: () => { if (!selected) choose(template); setOpen(false); } }, t('personal.useSame')) },
      loading ? h('span', { role: 'status' }, t('personal.processing'))
        : error ? h('div', { role: 'alert' }, h('p', null, error), h('button', { type: 'button', onClick: () => load(page) }, t('templates.retry')))
        : src ? h(FitPreviewImage, { src, alt: `${template.name} ${page + 1}`, zoom, onZoom: setZoom, onError: () => setError(t('personal.previewFailed')) })
        : h('span', { role: 'status' }, t('personal.previewFailed'))));
}
const FULL_TEMPLATE_PREVIEWS = /* GENERATED_PPT_FULL_PREVIEWS */ {};
const TEMPLATE_VIEWER_CSS = `
.ppt-preview-card {position:relative;min-width:0}
.ppt-preview-open {position:absolute;z-index:2;left:10px;top:10px;width:28px;height:28px;display:grid;place-items:center;padding:0;opacity:0;pointer-events:none;transition:opacity 120ms;border:1px solid var(--dsw-alias-border-l2-darkmode-thin);border-radius:7px;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font:inherit;font-size:12px;cursor:pointer}
.ppt-preview-card:hover .ppt-preview-open,.ppt-preview-card:focus-within .ppt-preview-open {opacity:1;pointer-events:auto}
@media (hover:none) {.ppt-preview-open {opacity:1;pointer-events:auto}}
@media (prefers-reduced-motion:reduce) {.ppt-preview-open {transition:none}}
.ppt-template-viewer {box-sizing:border-box;width:min(var(--dsh-chat-content-width,920px),calc(100vw - 32px));height:min(760px,calc(100dvh - 48px));max-width:none;max-height:none;padding:0;margin:auto;overflow:hidden;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-xl,20px);background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000014);color:var(--dsw-alias-label-primary,#222);font:inherit}
.ppt-template-viewer[open] {display:flex;flex-direction:column}
.ppt-template-viewer * {box-sizing:border-box}
.ppt-template-viewer::backdrop {background:#0007}
.ppt-template-viewer-header {display:flex;align-items:center;gap:12px;flex:none;padding:16px 20px;border-bottom:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd)}
.ppt-template-viewer h2 {flex:0 1 auto;min-width:0;margin:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:15px;font-weight:500;line-height:22px}
.ppt-template-page {flex:none;padding:2px 7px;border-radius:var(--dsw-radius-sm,6px);background:var(--dsw-alias-interactive-bg-hover,#f1f2f4);color:var(--dsw-alias-label-secondary,#666);font-size:12px;line-height:20px;font-variant-numeric:tabular-nums}
.ppt-template-viewer button {flex:none;min-width:32px;min-height:32px;padding:0;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));color:inherit;font:inherit;cursor:pointer}
.ppt-template-viewer button:hover:not(:disabled) {background:var(--dsw-alias-interactive-bg-hover,#f1f2f4)}
.ppt-template-viewer button:disabled {opacity:.35;cursor:default}
.ppt-template-viewer button:focus-visible,.ppt-preview-open:focus-visible {outline:2px solid var(--dsw-alias-state-business-primary,#3569d4);outline-offset:2px}
.ppt-template-viewer .ppt-template-close {margin-left:auto;border-color:transparent;color:var(--dsw-alias-label-secondary,#666);font-size:22px;line-height:1}
.ppt-template-canvas {position:relative;display:flex;flex:1;min-height:0;padding:24px 56px;background:var(--dsw-alias-bg-module-platform,var(--dsw-alias-interactive-bg-hover,#f1f2f4))}
.ppt-template-stage {position:relative;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;overflow:hidden}
.ppt-template-stage img {border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000014)}
.ppt-template-stage>[role] {display:grid;flex:1;place-content:center;margin:0;text-align:center;padding:16px;font-size:13px;overflow:auto}
.ppt-template-stage>[role] button {justify-self:center;padding:4px 12px}
.ppt-template-viewer .ppt-template-arrow {position:absolute;z-index:1;top:50%;width:32px;height:36px;transform:translateY(-50%);font-size:25px;line-height:1;box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000014)}
.ppt-template-previous {left:12px}.ppt-template-next {right:12px}
.ppt-template-viewer-footer {display:flex;align-items:center;flex:none;gap:16px;min-width:0;padding:12px 20px;border-top:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd)}
.ppt-template-thumbnails {display:flex;align-items:center;flex:1;min-width:0;gap:10px;padding:4px;overflow-x:auto;overscroll-behavior:contain;scrollbar-width:thin}
.ppt-template-viewer .ppt-template-thumb {position:relative;display:grid;place-items:center;width:104px;min-height:0;aspect-ratio:16/9;overflow:hidden;border-radius:var(--dsw-radius-sm,6px);background:var(--dsw-alias-bg-module-platform,var(--dsw-alias-interactive-bg-hover,#f1f2f4));color:var(--dsw-alias-label-secondary,#666)}
.ppt-template-thumb img {display:block;width:100%;height:100%;object-fit:contain}
.ppt-template-thumb span {position:absolute;right:3px;bottom:3px;min-width:17px;padding:0 3px;border-radius:3px;background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));font-size:10px;line-height:16px;font-variant-numeric:tabular-nums}
.ppt-template-viewer .ppt-template-thumb[aria-current=page] {border-color:var(--dsw-alias-state-business-primary,#3569d4);outline:2px solid color-mix(in srgb,var(--dsw-alias-state-business-primary,#3569d4) 24%,transparent);outline-offset:1px}
.ppt-template-viewer .ppt-preview-use {height:36px;padding:0 14px;border-color:transparent;background:var(--dsw-alias-label-primary,#222);color:var(--dsw-alias-label-primary-foreground,var(--dsw-alias-bg-base,#fff));font-size:13px;white-space:nowrap}
.ppt-template-viewer .ppt-preview-use:hover:not(:disabled) {background:color-mix(in srgb,var(--dsw-alias-label-primary,#222) 88%,var(--dsw-alias-bg-base,#fff))}
@media(max-width:600px) {.ppt-template-viewer{width:calc(100vw - 24px);height:calc(100dvh - 32px)}.ppt-template-viewer-header{padding:12px;gap:8px}.ppt-template-viewer h2{font-size:14px}.ppt-template-canvas{padding:16px 42px}.ppt-template-viewer .ppt-template-arrow{min-width:28px;width:28px}.ppt-template-previous{left:6px}.ppt-template-next{right:6px}.ppt-template-viewer-footer{flex-wrap:wrap;gap:8px;padding:10px 12px}.ppt-template-thumbnails{flex-basis:100%}.ppt-template-viewer .ppt-preview-use{margin-left:auto}.ppt-template-viewer .ppt-template-thumb{width:88px}}
`;

function PersonalTemplateManager({ client, mode, sessionId, state, choose, startAuthoring, t, mutable = true, builtInCards = [], showImport = false, showPersonal = true }) {
  const h = react.createElement;
  const [modal, setModal] = react.useState(null);
  const [name, setName] = react.useState('');
  const [description, setDescription] = react.useState('');
  const [operation, setOperation] = react.useState({ kind: 'idle' });
  const dialog = react.useRef(null);
  const fileInput = react.useRef(null);
  const [menu, setMenu] = react.useState(null);
  const menuRef = react.useRef(null);
  const alive = react.useRef(true);
  const titleId = react.useId();
  react.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, [sessionId]);
  react.useLayoutEffect(() => {
    if (!modal) return;
    const node = dialog.current;
    node.showModal(); node.querySelector('input,button')?.focus();
    return () => { if (node.open) node.close(); };
  }, [modal?.id, modal?.kind]);
  react.useEffect(() => {
    if (!menu) return;
    menuRef.current?.querySelector('[role=menuitem]')?.focus();
    const outside = event => { if (!event.target.closest('.personal-actions')) setMenu(null); };
    document.addEventListener('pointerdown', outside);
    return () => document.removeEventListener('pointerdown', outside);
  }, [menu]);
  const busy = operation.kind === 'busy';
  const templates = state.templates.filter(item => showPersonal && item.origin === 'personal').sort((a, b) => (b.createdAt ?? '').localeCompare(a.createdAt ?? ''));
  const close = () => { if (!busy) { setModal(null); setOperation({ kind: 'idle' }); } };
  const edit = template => { setName(template.name); setDescription(template.description ?? ''); setOperation({ kind: 'idle' }); setModal({ kind: 'edit', ...template }); };
  const act = async endpoint => {
    if (busy) return;
    setOperation({ kind: 'busy' });
    try {
      await client.call(endpoint, { templateId: modal.id, name, description });
      const next = await client.call('state');
      if (!alive.current) return;
      mode.setTemplateState(sessionId, next); setModal(null); setOperation({ kind: 'idle' });
    } catch (error) { if (alive.current) setOperation({ kind: 'error', message: error.message }); }
  };
  const transfer = async (action, value) => {
    if (busy || !mutable) return;
    setMenu(null); setOperation({ kind: 'busy', action });
    try {
      if (action === 'import') {
        if (!/\.dshtpl$/i.test(value.name) || value.size > 60 * 1024 * 1024 || !value.size) throw new Error(t('personal.packageFileHint'));
        const bytes = new Uint8Array(await value.arrayBuffer());
        const chunks = [];
        for (let offset = 0; offset < bytes.length; offset += 24576)
          chunks.push(btoa(String.fromCharCode(...bytes.subarray(offset, offset + 24576))));
        const result = await client.call('template/package-import', { fileName: value.name, base64: chunks.join('') });
        const next = await client.call('state');
        mode.setTemplateState(sessionId, next);
        if (alive.current) setOperation({ kind: 'success', message: t(result.outcome === 'existing' ? 'personal.packageExisting' : 'personal.packageImported').replace('{name}', result.template.name) });
      } else {
        const result = await client.call('template/package-export', { templateId: value.id });
        const chunks = [];
        for (let offset = 0; offset < result.base64.length; offset += 32768)
          chunks.push(Uint8Array.from(atob(result.base64.slice(offset, offset + 32768)), character => character.charCodeAt(0)));
        const url = URL.createObjectURL(new Blob(chunks, { type: 'application/zip' }));
        const link = document.createElement('a');
        link.href = url; link.download = result.fileName; document.body.appendChild(link);
        try { link.click(); } finally { link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
        if (alive.current) setOperation({ kind: 'success', message: t('personal.packageExported').replace('{name}', value.name) });
      }
    } catch (error) { if (alive.current) setOperation({ kind: 'error', message: error.message }); }
  };
  return h('div', { 'data-personal-template-library': '', style: { padding: '8px 4px 16px' } },
    h('style', null, PERSONAL_LIBRARY_CSS),
    mutable && showImport && h('div', { className: 'personal-library-toolbar' },
      h('button', { type: 'button', className: 'personal-import', disabled: busy || state.loading, onClick: () => fileInput.current.click() }, t(operation.kind === 'busy' && operation.action === 'import' ? 'personal.packageImporting' : 'personal.packageImport')),
      h('span', null, t('personal.packageFileHint')),
      h('input', { ref: fileInput, type: 'file', accept: '.dshtpl', hidden: true, 'aria-label': t('personal.packageImport'), disabled: busy,
        onChange: event => { const file = event.currentTarget.files?.[0]; event.currentTarget.value = ''; if (file) transfer('import', file); } })),
    !modal && operation.kind !== 'idle' && h('p', { className: 'personal-package-feedback', role: operation.kind === 'error' ? 'alert' : 'status' },
      operation.kind === 'busy' ? t(operation.action === 'import' ? 'personal.packageImporting' : 'personal.packageExporting') : operation.message),
    h('div', { className: OfficePptHero_module_css_default.templateGrid },
      mutable && showPersonal && h('button', { type: 'button', className: 'personal-create', onClick: () => startAuthoring(), disabled: busy || state.loading },
        h('span', { 'aria-hidden': true, style: { fontSize: 26 } }, '✦'), h('span', null, t('personal.create')), h('small', null, t('authoring.entryHint'))),
      ...templates.map(template => h('div', { key: template.id, 'data-personal-card': template.id },
        h('div', { className: 'personal-frame' },
          h(TemplatePreviewCard, { template, selected: template.id === state.selectedId, choose, client, t, disabled: busy || state.loading }),
          mutable && h('div', { className: 'personal-actions', onBlur: event => { if (!event.currentTarget.contains(event.relatedTarget)) setMenu(null); },
            onKeyDown: event => {
              if (event.key === 'Escape') { event.preventDefault(); setMenu(null); event.currentTarget.querySelector('button').focus(); }
              const items = [...event.currentTarget.querySelectorAll('[role=menuitem]')];
              if (items.length && ['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
                event.preventDefault(); const index = items.indexOf(document.activeElement);
                items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + items.length) % items.length]?.focus();
              }
            } },
            h('button', { type: 'button', className: 'personal-more', 'aria-label': `${t('personal.more')} · ${template.name}`, 'aria-haspopup': 'menu', 'aria-expanded': menu === template.id,
              disabled: busy || state.loading, onClick: () => setMenu(menu === template.id ? null : template.id) }, '⋯'),
            menu === template.id && h('div', { ref: menuRef, className: 'personal-menu', role: 'menu', 'aria-label': template.name },
              h('button', { type: 'button', role: 'menuitem', onClick: () => transfer('export', template) }, t('personal.packageExport')),
              h('button', { type: 'button', role: 'menuitem', onClick: () => { setMenu(null); startAuthoring(template.id); } }, t('personal.editContent')),
              h('button', { type: 'button', role: 'menuitem', 'aria-label': t('personal.edit'), onClick: () => { setMenu(null); edit(template); } }, t('personal.edit')),
              h('button', { type: 'button', role: 'menuitem', 'aria-label': t('personal.delete'), onClick: () => { setMenu(null); setModal({ ...template, kind: 'delete' }); setOperation({ kind: 'idle' }); } }, t('personal.delete'))))))),
      ...builtInCards.map(card => react.cloneElement(card, { disabled: busy || card.props.disabled }))),
    modal && h('dialog', { ref: dialog, className: 'personal-dialog', 'aria-labelledby': titleId,
      onCancel: event => { event.preventDefault(); close(); }, onKeyDown: event => {
        if (event.key !== 'Tab') return;
        const fields = [...event.currentTarget.querySelectorAll('button:not(:disabled),input:not(:disabled),textarea:not(:disabled)')];
        if (event.shiftKey && document.activeElement === fields[0]) { event.preventDefault(); fields.at(-1)?.focus(); }
        else if (!event.shiftKey && document.activeElement === fields.at(-1)) { event.preventDefault(); fields[0]?.focus(); }
      } },
      h('header', null, h('h2', { id: titleId }, t(modal.kind === 'delete' ? 'personal.deleteTitle' : 'personal.editTitle')),
        h('button', { type: 'button', 'aria-label': t('personal.close'), disabled: busy, onClick: close }, '×')),
      h('form', { onSubmit: event => { event.preventDefault(); act(modal.kind === 'delete' ? 'template/delete' : 'template/update'); } },
        modal.kind === 'delete' ? h('p', null, modal.name, ' ', t('personal.deleteHint')) : h(react.Fragment, null,
          h('label', null, t('personal.name'), h('input', { value: name, maxLength: 80, required: true, disabled: busy, onChange: event => setName(event.target.value) })),
          h('label', null, t('personal.description'), h('textarea', { value: description, maxLength: 1000, disabled: busy, onChange: event => setDescription(event.target.value) }))),
        operation.kind === 'error' && h('p', { role: 'alert' }, operation.message),
        h('footer', null,
          h('button', { type: 'button', disabled: busy, onClick: close }, t('personal.cancel')),
          h('button', { type: 'submit', className: modal.kind === 'delete' ? 'personal-danger' : 'personal-submit', disabled: busy || (modal.kind === 'edit' && !name.trim()) },
            t(busy ? 'personal.processing' : modal.kind === 'delete' ? 'personal.confirmDelete' : 'personal.update'))))));
}
const PERSONAL_LIBRARY_CSS = `
.personal-create{box-sizing:border-box;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;aspect-ratio:16/9;width:100%;padding:12px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:10px;background:var(--dsw-alias-interactive-bg-hover,#f1f2f4);color:var(--dsw-alias-label-secondary,#555);font:inherit;cursor:pointer}
.personal-create small{font-size:11px}.personal-create:hover{color:var(--dsw-alias-label-primary,#222)}
[data-personal-card]{min-width:0}.personal-frame{position:relative}.personal-actions{position:absolute;right:8px;top:8px;display:flex;gap:4px;z-index:2;opacity:0}.personal-frame:hover .personal-actions,.personal-frame:focus-within .personal-actions{opacity:1}
.personal-more{width:28px;height:28px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:7px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#222);font:inherit;font-size:20px;line-height:1;cursor:pointer}
.personal-actions:has([aria-expanded=true]){opacity:1;z-index:3}
.personal-menu{position:absolute;right:0;top:32px;display:grid;min-width:144px;padding:4px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-alias-bg-base,#fff);box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000020)}
.personal-menu button{padding:8px 12px;border:0;border-radius:6px;background:transparent;color:var(--dsw-alias-label-primary,#222);font:inherit;font-size:13px;text-align:left;white-space:nowrap;cursor:pointer}
.personal-menu button:hover,.personal-menu button:focus-visible{background:var(--dsw-alias-interactive-bg-hover,#f1f2f4)}
.personal-library-toolbar{display:flex;align-items:center;flex-wrap:wrap;gap:10px;margin:0 0 12px;color:var(--dsw-alias-label-secondary,#666);font-size:12px}
.personal-import{min-height:32px;padding:5px 12px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:transparent;color:var(--dsw-alias-label-primary,#222);font:inherit;font-size:13px;cursor:pointer}
.personal-import:hover{background:var(--dsw-alias-interactive-bg-hover,#f1f2f4)}
.personal-import:focus-visible,.personal-more:focus-visible,.personal-menu button:focus-visible{outline:2px solid var(--dsw-alias-state-business-primary,#3569d4);outline-offset:2px}
.personal-import:disabled,.personal-more:disabled,.personal-create:disabled{opacity:.5;cursor:default}
.personal-package-feedback{margin:0 0 12px;color:var(--dsw-alias-label-secondary,#666);font-size:13px;line-height:20px;overflow-wrap:anywhere}
.personal-package-feedback[role=alert]{color:var(--dsw-alias-state-error-primary,#b42318)}
.personal-dialog{box-sizing:border-box;width:min(560px,calc(100vw - 40px));max-height:calc(100dvh - 48px);margin:auto;padding:24px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:16px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#222);font:inherit;overflow-y:auto}
.personal-dialog::backdrop{background:#0007}.personal-dialog header,.personal-dialog footer{display:flex;justify-content:space-between;align-items:center;gap:12px}.personal-dialog h2{font-size:18px;margin:0}.personal-dialog form{display:grid;gap:18px;margin-top:20px}.personal-dialog label{display:grid;gap:8px}.personal-dialog input,.personal-dialog textarea{box-sizing:border-box;width:100%;padding:10px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ccc);border-radius:8px;background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit}.personal-dialog textarea{min-height:100px}.personal-dialog footer{justify-content:flex-end}.personal-dialog button{padding:7px 12px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ccc);border-radius:8px;background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit;cursor:pointer}.personal-dialog .personal-submit{background:var(--dsw-alias-label-primary,#222);color:var(--dsw-alias-bg-base,#fff)}.personal-dialog .personal-danger{color:#b42318}.personal-dialog button:disabled{opacity:.5}
@media(hover:none){.personal-actions{opacity:1}}
`;

function TemplateAuthoringPanel({ client, authoring, operation, running, save, exit, t }) {
  const h = react.createElement;
  const [name, setName] = react.useState(authoring.name ?? '');
  const [viewer, setViewer] = react.useState(false);
  const [page, setPage] = react.useState(1);
  const [preview, setPreview] = react.useState({ kind: 'loading' });
  const [retry, setRetry] = react.useState(0);
  const [zoom, setZoom] = react.useState(1);
  const titleId = react.useId();
  const busy = operation.kind === 'busy';
  const identity = `${authoring.workflowId}:${authoring.revision}:${authoring.draftId}`;
  react.useEffect(() => { setName(authoring.name ?? ''); setPage(1); setZoom(1); setViewer(false); }, [identity]);
  react.useEffect(() => {
    let alive = true;
    setPreview({ kind: 'loading' });
    client.call('template/authoring-preview', { workflowId: authoring.workflowId, revision: authoring.revision, draftId: authoring.draftId, page })
      .then(value => { if (alive) setPreview({ kind: 'ready', src: value.preview }); })
      .catch(error => { if (alive) setPreview({ kind: 'error', message: error.message }); });
    return () => { alive = false; };
  }, [client, identity, page, retry]);
  const navigate = index => { if (index + 1 !== page) setPreview({ kind: 'loading' }); setPage(index + 1); setZoom(1); };
  const previewContent = () => preview.kind === 'ready'
    ? h(FitPreviewImage, { src: preview.src, alt: `${t('personal.preview')} ${page}`, zoom, onZoom: setZoom, onError: () => setPreview({ kind: 'error', message: t('personal.previewFailed') }) })
    : preview.kind === 'error' ? h('div', { role: 'alert' }, h('p', null, preview.message), h('button', { type: 'button', className: 'dsh-ppt-authoring-button', onClick: () => setRetry(value => value + 1) }, t('templates.retry')))
    : h('p', { role: 'status' }, t('personal.processing'));
  return h('div', { className: 'dsh-ppt-authoring-frame' },
    h('section', { className: 'dsh-ppt-authoring', 'data-template-authoring': authoring.phase, 'aria-labelledby': titleId },
      h('style', null, TEMPLATE_VIEWER_CSS + AUTHORING_CSS),
      h('header', { className: 'dsh-ppt-authoring-header' },
        h('h2', { className: 'dsh-ppt-authoring-title', id: titleId }, t('authoring.title')),
        h('button', { type: 'button', className: 'dsh-ppt-authoring-button dsh-ppt-authoring-exit', disabled: busy, onClick: exit }, t('authoring.exit'))),
      h('div', { className: 'dsh-ppt-authoring-body' },
        operation.kind === 'error' && h('p', { className: 'dsh-ppt-authoring-notice', role: 'alert' }, operation.message),
        h('div', { className: 'dsh-ppt-authoring-review' },
          h('button', { type: 'button', className: 'dsh-ppt-authoring-thumbnail', onClick: () => setViewer(true), 'aria-label': t('authoring.preview') },
            preview.kind === 'ready' ? h('img', { src: preview.src, alt: `${t('personal.preview')} ${page}` }) : h('span', null, t('authoring.preview'))),
          h('div', { className: 'dsh-ppt-authoring-copy' },
            h('p', { className: 'dsh-ppt-authoring-summary' }, t('authoring.review').replace('{count}', authoring.slideCount)),
            h('small', { className: 'dsh-ppt-authoring-hint' }, t('authoring.revise'))),
          h('button', { type: 'button', className: 'dsh-ppt-authoring-button dsh-ppt-authoring-preview', onClick: () => setViewer(true) }, t('authoring.preview'))),
        preview.kind === 'error' && h('div', { className: 'dsh-ppt-authoring-notice' }, previewContent())),
      h('form', { className: 'dsh-ppt-authoring-footer', onSubmit: event => { event.preventDefault(); if (preview.kind === 'ready' && !busy && !running && name.trim()) save(name, authoring.origin?.kind === 'existing' ? 'original' : 'copy'); } },
        h('label', { className: 'dsh-ppt-authoring-name' },
          h('span', null, t('personal.name')),
          h('input', { className: 'dsh-ppt-authoring-input', value: name, maxLength: 80, required: true, disabled: busy || running, onChange: event => setName(event.target.value) })),
        h('div', { className: 'dsh-ppt-authoring-save-actions' },
          authoring.origin?.kind === 'existing' && h('button', { type: 'button', className: 'dsh-ppt-authoring-button', disabled: busy || running || preview.kind !== 'ready' || !name.trim(), onClick: () => save(name, 'copy') }, t('authoring.saveCopy')),
          h('button', { type: 'submit', className: 'dsh-ppt-authoring-button dsh-ppt-authoring-save', disabled: busy || running || preview.kind !== 'ready' || !name.trim() }, t(authoring.origin?.kind === 'existing' ? 'authoring.update' : 'authoring.save')))),
      viewer && h(TemplatePreviewViewer, { key: identity, title: authoring.name || t('authoring.preview'), page: page - 1, count: authoring.slideCount, navigate, close: () => setViewer(false), t, loading: preview.kind === 'loading',
        getThumbnail: index => index + 1 === page && preview.kind === 'ready' ? preview.src : undefined,
        loadThumbnail: async index => (await client.call('template/authoring-preview', { workflowId: authoring.workflowId, revision: authoring.revision, draftId: authoring.draftId, page: index + 1 })).preview }, previewContent())));
}
const AUTHORING_CSS = `
.dsh-ppt-authoring-frame{container:dsh-ppt-authoring / inline-size;box-sizing:border-box;display:flex;justify-content:center;flex:none;min-width:0;width:100%;padding:6px calc(var(--dsh-composer-side-clearance,16px) + 16px) 10px}
.dsh-ppt-authoring-feedback{box-sizing:border-box;display:flex;align-items:center;gap:10px;width:100%;max-width:var(--dsh-chat-content-width,920px);padding:12px 16px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));color:var(--dsw-alias-label-secondary,#666);font-size:13px;line-height:20px}
.dsh-ppt-authoring-feedback p{flex:1;min-width:0;margin:0;overflow-wrap:anywhere}
.dsh-ppt-authoring-success-icon{flex:none;color:var(--dsw-alias-state-success-primary,#16824b)}
.dsh-ppt-authoring-feedback .dsh-ppt-authoring-dismiss{width:28px;min-height:28px;padding:0;border-color:transparent;color:var(--dsw-alias-label-tertiary,#888);font-size:20px;line-height:1}
.dsh-ppt-authoring-error{box-sizing:border-box;width:100%;max-width:var(--dsh-chat-content-width,920px);padding:16px 24px;border-radius:var(--dsw-radius-xl,20px);background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000014);color:var(--dsw-alias-label-primary,#222);font-size:13px;line-height:20px;overflow-wrap:anywhere}
.dsh-ppt-authoring-error>p{margin:0 0 12px;color:var(--dsw-alias-state-error-primary,#b42318)}
.dsh-ppt-authoring-error>div{display:flex;flex-wrap:wrap;gap:8px}
.dsh-ppt-authoring{box-sizing:border-box;display:flex;flex-direction:column;min-width:0;width:100%;max-width:var(--dsh-chat-content-width,920px);max-height:min(60vh,520px);padding:0;border:0;border-radius:var(--dsw-radius-xl,20px);background:var(--dsw-specific-input-major,var(--dsw-alias-bg-base,#fff));box-shadow:var(--dsw-elevation-panel,0 2px 12px #00000014);color:var(--dsw-alias-label-primary,#222);font:inherit;font-size:14px;line-height:1.5;overflow:hidden;--dsw-elevation-stroke-color:var(--dsw-alias-border-l2-darkmode-thin,#ddd);--dsh-scrollbar-thumb:var(--dsw-alias-scrollbar-bg-l2,#ccc);--dsh-scrollbar-thumb-hover:var(--dsw-alias-scrollbar-hover-l2,#aaa)}
.dsh-ppt-authoring *{box-sizing:border-box}
.dsh-ppt-authoring-header{display:flex;justify-content:space-between;align-items:center;flex:none;gap:16px;padding:16px 16px 14px 24px;border-bottom:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd)}
.dsh-ppt-authoring-title{min-width:0;margin:0;font-size:16px;font-weight:500;line-height:22px;overflow-wrap:anywhere}
.dsh-ppt-authoring-body{flex:1;min-height:0;padding:16px 24px;overflow-y:auto;overscroll-behavior:contain}
.dsh-ppt-authoring-review{display:flex;align-items:center;gap:16px;padding:16px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-alias-bg-module-platform,var(--dsw-alias-interactive-bg-hover,#f1f2f4))}
.dsh-ppt-authoring-thumbnail{display:grid;place-items:center;flex:none;width:144px;aspect-ratio:16/9;padding:4px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-alias-bg-module-platform,#f5f5f5);color:var(--dsw-alias-label-secondary,#666);font:inherit;font-size:12px;cursor:pointer;overflow:hidden}
.dsh-ppt-authoring-thumbnail img{display:block;width:100%;height:100%;object-fit:contain}
.dsh-ppt-authoring-copy{display:flex;flex-direction:column;align-items:flex-start;flex:1;min-width:0;gap:6px}
.dsh-ppt-authoring-summary{margin:0;font-size:14px;line-height:22px;overflow-wrap:anywhere}
.dsh-ppt-authoring-hint{display:block;color:var(--dsw-alias-label-secondary,#666);font-size:12px;line-height:20px;overflow-wrap:anywhere}
.dsh-ppt-authoring-button{display:inline-flex;justify-content:center;align-items:center;flex:none;min-height:32px;padding:5px 12px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:var(--dsw-radius-md,10px);background:transparent;color:inherit;font:inherit;font-size:13px;line-height:20px;white-space:nowrap;cursor:pointer;transition:background-color .12s,border-color .12s}
.dsh-ppt-authoring-button:hover:not(:disabled),.dsh-ppt-authoring-thumbnail:hover:not(:disabled){border-color:var(--dsw-alias-border-l4,#aaa);background:var(--dsw-alias-interactive-bg-hover,#f1f2f4)}
.dsh-ppt-authoring-exit{min-height:28px;padding:3px 8px;border-color:transparent;color:var(--dsw-alias-label-secondary,#666);font-size:12px}
.dsh-ppt-authoring-button:disabled{opacity:.4;cursor:default}
.dsh-ppt-authoring-button:focus-visible,.dsh-ppt-authoring-thumbnail:focus-visible,.dsh-ppt-authoring-input:focus-visible{outline:2px solid var(--dsw-alias-state-business-primary,#3569d4);outline-offset:2px}
.dsh-ppt-authoring-footer{display:flex;align-items:center;flex:none;gap:12px;margin:0;padding:14px 24px;background:color-mix(in srgb,var(--dsw-alias-interactive-bg-hover,#f1f2f4) 45%,transparent);border-top:1px solid var(--dsw-alias-border-l2-darkmode-thin,#e5e5e5)}
.dsh-ppt-authoring-save-actions{display:flex;flex:none;gap:8px}
.dsh-ppt-authoring-name{display:flex;align-items:center;flex:1;min-width:0;gap:12px;font-size:13px;line-height:20px}
.dsh-ppt-authoring-name>span{flex:none;white-space:nowrap;font-weight:500}
.dsh-ppt-authoring-input{flex:1;min-width:0;width:100%;height:36px;padding:7px 12px;border:1px solid var(--dsw-alias-border-l4,#ccc);border-radius:var(--dsw-radius-md,10px);background:var(--dsw-alias-bg-module-platform,var(--dsw-alias-bg-base,#fff));color:inherit;caret-color:var(--dsw-alias-state-business-primary,#3569d4);font:inherit}
.dsh-ppt-authoring-input:hover:not(:disabled){border-color:var(--dsw-alias-label-tertiary,#999)}
.dsh-ppt-authoring-input:disabled{opacity:.5;cursor:default}
.dsh-ppt-authoring-save{height:36px;min-height:36px;padding:7px 16px;border-color:transparent;background:var(--dsw-alias-label-primary,#222);color:var(--dsw-alias-label-primary-foreground,var(--dsw-alias-bg-base,#fff));font-weight:500}
.dsh-ppt-authoring-save:hover:not(:disabled){border-color:transparent;background:color-mix(in srgb,var(--dsw-alias-label-primary,#222) 88%,var(--dsw-alias-bg-base,#fff))}
.dsh-ppt-authoring-notice{margin:0 0 12px;color:var(--dsw-alias-state-error-primary,#b42318);font-size:13px;line-height:20px;overflow-wrap:anywhere}
.dsh-ppt-authoring-notice:last-child{margin:12px 0 0}
.dsh-ppt-authoring-notice p{margin:0 0 8px}
@container dsh-ppt-authoring (max-width:656px){.dsh-ppt-authoring-header{padding:14px 12px 12px 18px}.dsh-ppt-authoring-title{font-size:15px;line-height:21px}.dsh-ppt-authoring-body{padding:14px 18px}.dsh-ppt-authoring-footer{padding:14px 18px}.dsh-ppt-authoring-review{display:grid;grid-template-columns:112px minmax(0,1fr);gap:12px}.dsh-ppt-authoring-thumbnail{grid-row:1/3;width:112px}.dsh-ppt-authoring-preview{grid-column:2;justify-self:start}}
@container dsh-ppt-authoring (max-width:536px){.dsh-ppt-authoring-footer{flex-direction:column;align-items:stretch}.dsh-ppt-authoring-name{flex:none}.dsh-ppt-authoring-save-actions{width:100%}.dsh-ppt-authoring-save-actions>button{flex:1}.dsh-ppt-authoring-review{padding:12px;grid-template-columns:96px minmax(0,1fr)}.dsh-ppt-authoring-thumbnail{width:96px}}
@container dsh-ppt-authoring (max-width:400px){.dsh-ppt-authoring-name{flex-direction:column;align-items:stretch;gap:6px}.dsh-ppt-authoring-review{grid-template-columns:1fr}.dsh-ppt-authoring-thumbnail{grid-row:auto;width:100%;max-width:192px}.dsh-ppt-authoring-preview{grid-column:auto}}
@media(prefers-reduced-motion:reduce){.dsh-ppt-authoring-button{transition:none}}
`;
