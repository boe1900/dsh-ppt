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

/** Shared detailed template viewer; selection remains an independent action. */
function TemplatePreviewCard({ template, selected, choose, client, t, disabled }) {
  const h = react.createElement;
  const [open, setOpen] = react.useState(false);
  const [page, setPage] = react.useState(0);
  const [pages, setPages] = react.useState({});
  const [loading, setLoading] = react.useState(false);
  const [error, setError] = react.useState('');
  const [zoom, setZoom] = react.useState(1);
  const dialog = react.useRef(null);
  const request = react.useRef(0);
  const titleId = react.useId();
  const staticPages = FULL_TEMPLATE_PREVIEWS[template.id] ?? template.previewImages ?? [];
  const detailedPersonal = template.origin === 'personal' && client.bound === true;
  const slides=personalPreviewSlides(template);
  const count = detailedPersonal ? slides.length : staticPages.length;
  const cacheKey=slide=>`${template.previewVersion ?? 0}:${slide}`;
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
  react.useLayoutEffect(() => {
    if (!open) return;
    const node = dialog.current;
    node.showModal();
    node.querySelector('button')?.focus();
    load(0);
    return () => { request.current++; if (node.open) node.close(); };
  }, [open,template.previewVersion]);
  react.useEffect(() => () => { request.current++; }, [client, template.id,template.previewVersion]);
  return h('div', { className: 'ppt-preview-card' },
    h('style', null, TEMPLATE_VIEWER_CSS),
    h(TemplateCard, { template, selected, choose, disabled }),
    h('button', { type: 'button', className: 'ppt-preview-open', 'aria-label': `${t('personal.view')} · ${template.name}`, title: t('personal.view'), onClick: () => setOpen(true) },
      h('svg', { width: 16, height: 16, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.6, 'aria-hidden': true },
        h('path', { d: 'M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z' }),
        h('circle', { cx: 12, cy: 12, r: 3 }))),
    open && h('dialog', { ref: dialog, className: 'ppt-template-viewer', 'aria-labelledby': titleId,
      onCancel: event => { event.preventDefault(); setOpen(false); },
      onClick: event => {
        if (event.target !== event.currentTarget) return;
        const rect = event.currentTarget.getBoundingClientRect();
        if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) setOpen(false);
      },
      onKeyDown: event => {
        if (event.key === 'ArrowLeft' && page > 0) { event.preventDefault(); load(page - 1); }
        if (event.key === 'ArrowRight' && page < count - 1) { event.preventDefault(); load(page + 1); }
        if (event.key === 'Tab') {
          const buttons = [...event.currentTarget.querySelectorAll('button:not(:disabled)')];
          const first = buttons[0], last = buttons.at(-1);
          if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
          else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
        }
      } },
      h('header', null, h('h2', { id: titleId }, template.name), h('button', { type: 'button', 'aria-label': t('personal.close'), onClick: () => setOpen(false) }, '×')),
      h('div', { className: 'ppt-template-stage', 'aria-busy': loading },
        loading ? h('span', { role: 'status' }, t('personal.processing'))
          : error ? h('div', { role: 'alert' }, h('p', null, error), h('button', { type: 'button', onClick: () => load(page) }, t('templates.retry')))
          : src ? h(FitPreviewImage, { src, alt: `${template.name} ${page + 1}`, zoom, onZoom: setZoom, onError: () => setError(t('personal.previewFailed')) })
          : h('span', { role: 'status' }, t('personal.previewFailed'))),
      h('footer', null,
        h('button', { type: 'button', disabled: page === 0, 'aria-label': t('personal.previous'), onClick: () => load(page - 1) }, '‹'),
        h('span', { role: 'status' }, `${template.templateSamples ? t('personal.layouts')+' ' : ''}${page + 1} / ${count}`),
        h('button', { type: 'button', disabled: page >= count - 1, 'aria-label': t('personal.next'), onClick: () => load(page + 1) }, '›'),
        h('button',{type:'button',className:'ppt-preview-use',disabled,onClick:()=>{if(!selected)choose(template);setOpen(false);}},t('personal.useSame')))));
}
const FULL_TEMPLATE_PREVIEWS = /* GENERATED_PPT_FULL_PREVIEWS */ {};
const TEMPLATE_VIEWER_CSS = `
.ppt-preview-card {position:relative;min-width:0}
.ppt-preview-open {position:absolute;z-index:2;left:10px;top:10px;width:28px;height:28px;display:grid;place-items:center;padding:0;opacity:0;pointer-events:none;transition:opacity 120ms;border:1px solid var(--dsw-alias-border-l2-darkmode-thin);border-radius:7px;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font:inherit;font-size:12px;cursor:pointer}
.ppt-preview-card:hover .ppt-preview-open,.ppt-preview-card:focus-within .ppt-preview-open {opacity:1;pointer-events:auto}
@media (hover:none) {.ppt-preview-open {opacity:1;pointer-events:auto}}
@media (prefers-reduced-motion:reduce) {.ppt-preview-open {transition:none}}
.ppt-template-viewer {box-sizing:border-box;width:min(1280px,calc(100vw - 32px));height:min(900px,calc(100dvh - 32px));padding:20px;margin:auto;border:1px solid var(--dsw-alias-border-l2-darkmode-thin);border-radius:16px;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font:inherit}
.ppt-template-viewer[open] {display:flex;flex-direction:column;gap:16px}
.ppt-template-viewer::backdrop {background:color-mix(in srgb,var(--dsw-alias-label-primary) 45%,transparent)}
.ppt-template-viewer header,.ppt-template-viewer footer {display:flex;align-items:center;gap:12px;flex:none}
.ppt-template-viewer .ppt-preview-use {width:auto;padding:0 14px;margin-left:8px;background:var(--dsw-alias-label-primary);color:var(--dsw-alias-bg-base);border-color:transparent;font-size:13px}
.ppt-template-viewer header {justify-content:space-between}
.ppt-template-viewer h2 {font-size:16px;margin:0;overflow-wrap:anywhere}
.ppt-template-viewer footer {justify-content:center;flex-wrap:wrap;font-size:13px;font-variant-numeric:tabular-nums}
.ppt-template-viewer button {min-width:32px;min-height:32px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin);border-radius:7px;background:var(--dsw-alias-bg-base);color:inherit;font:inherit;cursor:pointer}
.ppt-template-viewer button:disabled {opacity:.35;cursor:default}
.ppt-template-viewer button:focus-visible,.ppt-preview-open:focus-visible {outline:2px solid var(--dsw-alias-state-business-primary);outline-offset:2px}
.ppt-template-stage {display:flex;flex-direction:column;flex:1;min-height:0;overflow:auto;border-radius:8px;background:var(--dsw-alias-interactive-bg-hover);position:relative}
.ppt-template-stage img {display:block;height:auto;margin:0 auto}
.ppt-template-stage>[role] {display:grid;place-content:center;min-height:240px;text-align:center;padding:16px}
`;

function PersonalTemplateManager({ client, mode, sessionId, state, choose, startAuthoring, t, mutable = true, builtInCards = [] }) {
  const h = react.createElement;
  const [modal, setModal] = react.useState(null);
  const [name, setName] = react.useState('');
  const [description, setDescription] = react.useState('');
  const [operation, setOperation] = react.useState({ kind: 'idle' });
  const dialog = react.useRef(null);
  const alive = react.useRef(true);
  const titleId = react.useId();
  react.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, [sessionId]);
  react.useLayoutEffect(() => {
    if (!modal) return;
    const node = dialog.current;
    node.showModal(); node.querySelector('input,button')?.focus();
    return () => { if (node.open) node.close(); };
  }, [modal?.id, modal?.kind]);
  const busy = operation.kind === 'busy';
  const templates = state.templates.filter(item => item.origin === 'personal').sort((a, b) => (b.createdAt ?? '').localeCompare(a.createdAt ?? ''));
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
  return h('div', { 'data-personal-template-library': '', style: { padding: '8px 4px 16px' } },
    h('style', null, PERSONAL_LIBRARY_CSS),
    h('div', { className: OfficePptHero_module_css_default.templateGrid },
      mutable && h('button', { type: 'button', className: 'personal-create', onClick: startAuthoring, disabled: state.loading },
        h('span', { 'aria-hidden': true, style: { fontSize: 26 } }, '✦'), h('span', null, t('personal.create')), h('small', null, t('authoring.entryHint'))),
      ...templates.map(template => h('div', { key: template.id, 'data-personal-card': template.id },
        h('div', { className: 'personal-frame' },
          h(TemplatePreviewCard, { template, selected: template.id === state.selectedId, choose, client, t, disabled: state.loading }),
          mutable && h('div', { className: 'personal-actions' },
            h('button', { type: 'button', 'aria-label': t('personal.edit'), onClick: () => edit(template) }, '✎'),
            h('button', { type: 'button', 'aria-label': t('personal.delete'), onClick: () => { setModal({ ...template, kind: 'delete' }); setOperation({ kind: 'idle' }); } }, '×'))))),
      ...builtInCards),
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
.personal-actions button{width:28px;height:28px;border:1px solid #ffffff26;border-radius:7px;background:#222c;color:#fff;cursor:pointer}
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
  const dialog = react.useRef(null);
  const titleId = react.useId();
  const busy = operation.kind === 'busy';
  const identity = `${authoring.workflowId}:${authoring.revision}`;
  react.useEffect(() => { setName(authoring.name ?? ''); setPage(1); setZoom(1); setViewer(false); }, [identity]);
  react.useEffect(() => {
    let alive = true;
    setPreview({ kind: 'loading' });
    client.call('template/authoring-preview', { workflowId: authoring.workflowId, revision: authoring.revision, draftId: authoring.draftId, page })
      .then(value => { if (alive) setPreview({ kind: 'ready', src: value.preview }); })
      .catch(error => { if (alive) setPreview({ kind: 'error', message: error.message }); });
    return () => { alive = false; };
  }, [client, identity, page, retry]);
  react.useLayoutEffect(() => {
    if (!viewer) return;
    const node = dialog.current;
    node.showModal(); node.querySelector('button')?.focus();
    return () => { if (node.open) node.close(); };
  }, [viewer]);
  const previewContent = () => preview.kind === 'ready'
    ? h(FitPreviewImage, { src: preview.src, alt: `${t('personal.preview')} ${page}`, zoom, onZoom: setZoom, onError: () => setPreview({ kind: 'error', message: t('personal.previewFailed') }) })
    : preview.kind === 'error' ? h('div', { role: 'alert' }, h('p', null, preview.message), h('button', { type: 'button', onClick: () => setRetry(value => value + 1) }, t('templates.retry')))
    : h('p', { role: 'status' }, t('personal.processing'));
  return h('section', { className: 'dsh-ppt-authoring', 'data-template-authoring': authoring.phase, 'aria-labelledby': titleId },
    h('style', null, TEMPLATE_VIEWER_CSS + AUTHORING_CSS),
    h('header', null, h('strong', { id: titleId }, t('authoring.title')), h('button', { type: 'button', disabled: busy, onClick: exit }, t('authoring.exit'))),
    operation.kind === 'error' && h('p', { role: 'alert' }, operation.message),
    h('div', { className: 'dsh-ppt-authoring-review' },
      h('button', { type: 'button', className: 'dsh-ppt-authoring-thumbnail', onClick: () => setViewer(true), 'aria-label': t('authoring.preview') },
        preview.kind === 'ready' ? h('img', { src: preview.src, alt: `${t('personal.preview')} ${page}` }) : h('span', null, t('authoring.preview'))),
      h('div', null, h('p', null, t('authoring.review').replace('{count}', authoring.slideCount)),
        h('button', { type: 'button', onClick: () => setViewer(true) }, t('authoring.preview')), h('small', null, t('authoring.revise')))),
    h('form', { onSubmit: event => { event.preventDefault(); if (preview.kind === 'ready' && !running && name.trim()) save(name); } },
      h('label', null, t('personal.name'), h('input', { value: name, maxLength: 80, required: true, disabled: busy || running, onChange: event => setName(event.target.value) })),
      h('button', { type: 'submit', disabled: busy || running || preview.kind !== 'ready' || !name.trim() }, t('authoring.save'))),
    preview.kind === 'error' && previewContent(),
    viewer && h('dialog', { ref: dialog, className: 'ppt-template-viewer', 'aria-label': t('authoring.preview'),
      onCancel: event => { event.preventDefault(); setViewer(false); }, onKeyDown: event => {
        if (event.key === 'ArrowLeft' && page > 1) { event.preventDefault(); setPage(value => value - 1); }
        if (event.key === 'ArrowRight' && page < authoring.slideCount) { event.preventDefault(); setPage(value => value + 1); }
        if (event.key === 'Tab') {
          const fields = [...event.currentTarget.querySelectorAll('button:not(:disabled)')];
          if (event.shiftKey && document.activeElement === fields[0]) { event.preventDefault(); fields.at(-1)?.focus(); }
          else if (!event.shiftKey && document.activeElement === fields.at(-1)) { event.preventDefault(); fields[0]?.focus(); }
        }
      } },
      h('header', null, h('h2', null, authoring.name), h('button', { type: 'button', 'aria-label': t('personal.close'), onClick: () => setViewer(false) }, '×')),
      h('div', { className: 'ppt-template-stage' }, previewContent()),
      h('footer', null,
        h('button', { type: 'button', disabled: page <= 1, 'aria-label': t('personal.previous'), onClick: () => { setPage(value => value - 1); setZoom(1); } }, '‹'),
        h('span', null, `${page} / ${authoring.slideCount}`),
        h('button', { type: 'button', disabled: page >= authoring.slideCount, 'aria-label': t('personal.next'), onClick: () => { setPage(value => value + 1); setZoom(1); } }, '›'),
        h('span', null, t('authoring.revise')))));
}
const AUTHORING_CSS = `
.dsh-ppt-authoring{box-sizing:border-box;align-self:center;width:calc(100% - 2 * var(--dsh-composer-side-clearance,16px));padding:12px 16px;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ddd);border-radius:12px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#222);font:13px/1.5 var(--font-sans,system-ui,sans-serif)}
.dsh-ppt-authoring header,.dsh-ppt-authoring form,.dsh-ppt-authoring-review{display:flex;align-items:center;gap:12px}.dsh-ppt-authoring header{justify-content:space-between}.dsh-ppt-authoring p{margin:8px 0}.dsh-ppt-authoring small{display:block;color:var(--dsw-alias-label-secondary,#666)}.dsh-ppt-authoring button{border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ccc);border-radius:7px;padding:5px 10px;background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit;cursor:pointer}.dsh-ppt-authoring button:disabled{opacity:.5;cursor:default}.dsh-ppt-authoring form{margin-top:12px;flex-wrap:wrap}.dsh-ppt-authoring label{display:flex;align-items:center;gap:8px;flex:1}.dsh-ppt-authoring input{min-width:80px;width:100%;padding:6px 10px;box-sizing:border-box;border:1px solid var(--dsw-alias-border-l2-darkmode-thin,#ccc);border-radius:7px;background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit}.dsh-ppt-authoring form button{background:var(--dsw-alias-label-primary,#222);color:var(--dsw-alias-bg-base,#fff)}.dsh-ppt-authoring-review{margin-top:10px}.dsh-ppt-authoring .dsh-ppt-authoring-thumbnail{display:grid;place-items:center;flex:none;width:130px;aspect-ratio:16/9;padding:0;overflow:hidden}.dsh-ppt-authoring-thumbnail img{width:100%;height:100%;object-fit:contain}.dsh-ppt-authoring [role=alert]{color:var(--dsw-alias-state-error-primary,#b42318)}
@media(max-width:500px){.dsh-ppt-authoring-review{align-items:flex-start}.dsh-ppt-authoring .dsh-ppt-authoring-thumbnail{width:90px}}
`;
