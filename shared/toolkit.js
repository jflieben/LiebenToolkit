// LC Toolkit shell: theme, header with tool menu, related tools, footer, toasts,
// dialogs and small DOM helpers. Load in <head> without defer so the theme is set
// before first paint; the shell renders on DOMContentLoaded.
//
// Page contract:
//   <body data-tool="InactiveGuests">       tool id from registry.js
//   <div data-tk="header"></div>            replaced by the header
//   <div data-tk="related"></div>           optional, related tools
//   <div data-tk="footer"></div>            replaced by the footer
//   <div data-tk="about"></div>             optional, standard privacy/safety block
//   <div data-tk="debug"></div>             optional, standard debug log card
(() => {
  const THEME_KEY = 'lctk-theme';
  const R = () => window.ToolkitRegistry;

  // Where /tools/ lives relative to this page: tool pages sit one folder down.
  const scriptSrc = document.currentScript ? document.currentScript.src : '';
  const SHARED = scriptSrc ? scriptSrc.replace(/[^/]*$/, '') : '../shared/';
  const ROOT = SHARED.replace(/shared\/$/, '');
  const ICONS = SHARED + 'icons.svg';

  // ---------- theme ----------
  function storedTheme() {
    try { return localStorage.getItem(THEME_KEY); } catch { return null; }
  }
  function applyTheme(theme) {
    if (theme === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    else document.documentElement.removeAttribute('data-theme');
  }
  function currentTheme() { return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
  function setTheme(theme) {
    applyTheme(theme);
    try { localStorage.setItem(THEME_KEY, theme); } catch { /* storage blocked */ }
    document.querySelectorAll('[data-tk-theme-icon]').forEach((u) => u.setAttribute('href', `${ICONS}#${theme === 'dark' ? 'i-sun' : 'i-moon'}`));
    document.dispatchEvent(new CustomEvent('toolkit:theme', { detail: { theme } }));
  }
  // Explicit preference wins; a ?theme= query is for previews and screenshots.
  const queryTheme = new URLSearchParams(location.search).get('theme');
  applyTheme(queryTheme === 'dark' || queryTheme === 'light' ? queryTheme : storedTheme());

  // ---------- DOM helpers ----------
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  function esc(value) { return String(value == null ? '' : value).replace(/[&<>"']/g, (c) => ESC[c]); }

  // Only http(s), mailto and relative links survive; everything else becomes '#'.
  function safeUrl(value) {
    const s = String(value == null ? '' : value).trim();
    if (!s) return '#';
    try {
      const u = new URL(s, location.href);
      return ['http:', 'https:', 'mailto:'].includes(u.protocol) ? u.href : '#';
    } catch { return '#'; }
  }

  // h('div.card#id', { text, html, class, attrs..., on: { click } }, [children])
  function h(spec, props, children) {
    const m = /^([a-z0-9-]+)?((?:[.#][\w-]+)*)$/i.exec(spec) || [];
    const el = document.createElement(m[1] || 'div');
    (m[2] || '').replace(/([.#])([\w-]+)/g, (_, kind, name) => {
      if (kind === '.') el.classList.add(name); else el.id = name;
      return '';
    });
    for (const [k, v] of Object.entries(props || {})) {
      if (v == null || v === false) continue;
      if (k === 'text') el.textContent = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k === 'class') String(v).split(/\s+/).filter(Boolean).forEach((c) => el.classList.add(c));
      else if (k === 'on') Object.entries(v).forEach(([ev, fn]) => el.addEventListener(ev, fn));
      else if (k === 'href' || k === 'src') el.setAttribute(k, safeUrl(v));
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else el.setAttribute(k, v === true ? '' : v);
    }
    const kids = Array.isArray(children) ? children : children == null ? [] : [children];
    for (const c of kids.flat()) {
      if (c == null || c === false) continue;
      el.append(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return el;
  }

  const SVG_NS = 'http://www.w3.org/2000/svg';
  function icon(name, cls) {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('class', `icon${cls ? ' ' + cls : ''}`);
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');
    const use = document.createElementNS(SVG_NS, 'use');
    use.setAttribute('href', `${ICONS}#${name}`);
    svg.appendChild(use);
    return svg;
  }
  // Same icon as markup, for code that still builds HTML strings.
  function iconHtml(name, cls) {
    return `<svg class="icon${cls ? ' ' + esc(cls) : ''}" aria-hidden="true" focusable="false"><use href="${esc(ICONS)}#${esc(name)}"></use></svg>`;
  }

  // Replace <i data-icon="i-name"> placeholders in static markup with real icons.
  function hydrateIcons(root) {
    (root || document).querySelectorAll('i[data-icon]').forEach((i) => {
      const svg = icon(i.dataset.icon, i.className || '');
      i.replaceWith(svg);
    });
  }

  // ---------- toasts ----------
  let toastHost = null;
  const TOAST_ICONS = { info: 'i-info', ok: 'i-check-circle', warn: 'i-alert', err: 'i-error' };
  function toast(message, kind = 'info', { timeout } = {}) {
    if (kind === 'error' || kind === 'danger') kind = 'err';
    if (kind === 'success') kind = 'ok';
    if (!toastHost) {
      toastHost = h('div.tk-toasts', { role: 'status', 'aria-live': 'polite' });
      document.body.appendChild(toastHost);
    }
    const close = h('button.tk-toast-close', { type: 'button', 'aria-label': 'Dismiss' }, icon('i-close', 'sm'));
    const el = h('div.tk-toast', { class: kind }, [icon(TOAST_ICONS[kind] || 'i-info'), h('span', { text: message }), close]);
    const dismiss = () => {
      el.classList.add('leaving');
      setTimeout(() => el.remove(), 200);
    };
    close.addEventListener('click', dismiss);
    toastHost.appendChild(el);
    while (toastHost.children.length > 4) toastHost.firstChild.remove();
    setTimeout(dismiss, timeout || (kind === 'err' ? 8000 : 4500));
    return el;
  }

  // ---------- dialogs ----------
  // modal({ title, body: Node|string(text), html: trusted string, actions: [{ label, value, variant, icon }], wide, danger })
  // resolves with the clicked action's value, or null when dismissed.
  // The returned promise has .close() to dismiss the dialog from code.
  function modal({ title, body, html, actions, wide, danger, dismissable = true } = {}) {
    let closeFn = null;
    const promise = new Promise((resolve) => {
      const previous = document.activeElement;
      const bodyEl = h('div.tk-modal-body');
      if (body instanceof Node) bodyEl.appendChild(body);
      else if (html != null) bodyEl.innerHTML = html;
      else if (body != null) String(body).split(/\n{2,}/).forEach((p) => bodyEl.appendChild(h('p', { text: p })));

      const buttons = (actions || [{ label: 'Close', value: null }]).map((a) => {
        const cls = a.variant === 'primary' ? 'btn btn-primary' : a.variant === 'danger' ? 'btn btn-danger solid' : 'btn';
        return h('button', { type: 'button', class: cls, on: { click: () => done(a.value) } }, [a.icon ? icon(a.icon) : null, a.label]);
      });
      const titleId = `tk-modal-title-${Date.now()}`;
      const closeBtn = dismissable ? h('button.btn-icon.borderless', { type: 'button', 'aria-label': 'Close', on: { click: () => done(null) } }, icon('i-close')) : null;
      const dialog = h('div.tk-modal', { class: `${wide ? 'wide' : ''} ${danger ? 'danger' : ''}`, role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': titleId }, [
        h('div.tk-modal-head', null, [h('h2', { id: titleId, text: title || '' }), closeBtn]),
        bodyEl,
        h('div.tk-modal-foot', null, buttons),
      ]);
      const host = h('div.tk-modal-host', null, dialog);
      host.addEventListener('mousedown', (e) => { if (dismissable && e.target === host) done(null); });
      const onKey = (e) => {
        if (e.key === 'Escape' && dismissable) { e.preventDefault(); done(null); }
        if (e.key === 'Tab') {
          const f = dialog.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
          if (!f.length) return;
          const first = f[0], last = f[f.length - 1];
          if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
          else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
        }
      };
      document.addEventListener('keydown', onKey, true);
      function done(value) {
        document.removeEventListener('keydown', onKey, true);
        host.remove();
        if (previous && previous.focus) try { previous.focus(); } catch { /* gone */ }
        resolve(value);
      }
      document.body.appendChild(host);
      closeFn = () => { if (host.isConnected) done(null); };
      // Focus the safe choice, not the destructive one.
      (buttons.find((b) => !b.classList.contains('btn-danger') && !b.classList.contains('btn-primary')) || buttons[0] || dialog).focus();
    });
    promise.close = () => closeFn && closeFn();
    return promise;
  }

  // confirm({ title, body|html, okLabel, danger, items: [string] }) -> Promise<boolean>
  function confirm({ title, body, html, okLabel = 'Confirm', cancelLabel = 'Cancel', danger = false, items, itemsMax = 25, wide } = {}) {
    let content = body;
    if (items && items.length) {
      const wrap = h('div');
      if (body instanceof Node) wrap.appendChild(body);
      else if (html != null) wrap.appendChild(h('div', { html }));
      else if (body) wrap.appendChild(h('p', { text: body }));
      // An item is a string, or { text, tone: 'warn' | 'danger' } to flag it.
      wrap.appendChild(h('ul.tk-modal-list', null, items.slice(0, itemsMax).map((t) => (typeof t === 'string' ? h('li', { text: t }) : h('li', { class: t.tone || '', text: t.text })))));
      if (items.length > itemsMax) wrap.appendChild(h('p.muted.small', { text: `+ ${items.length - itemsMax} more` }));
      content = wrap;
      html = null;
    }
    return modal({
      title, body: content, html, wide, danger,
      actions: [
        { label: cancelLabel, value: false },
        { label: okLabel, value: true, variant: danger ? 'danger' : 'primary' },
      ],
    }).then((v) => v === true);
  }

  // Disable a button and show a spinner while fn runs.
  async function busy(btn, fn) {
    if (!btn) return fn();
    const was = btn.disabled;
    btn.disabled = true;
    btn.classList.add('is-busy');
    btn.setAttribute('aria-busy', 'true');
    try { return await fn(); }
    finally {
      btn.disabled = was;
      btn.classList.remove('is-busy');
      btn.removeAttribute('aria-busy');
    }
  }

  // Wire .tab[data-tab] buttons to .tab-panel#tab-<id>. Returns activate(id).
  function tabs(nav, { onChange } = {}) {
    const navEl = typeof nav === 'string' ? document.querySelector(nav) : nav;
    if (!navEl) return () => {};
    navEl.setAttribute('role', 'tablist');
    const buttons = [...navEl.querySelectorAll('[data-tab]')];
    function activate(id) {
      buttons.forEach((b) => {
        const on = b.dataset.tab === id;
        b.classList.toggle('active', on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
        b.tabIndex = on ? 0 : -1;
      });
      document.querySelectorAll('.tab-panel').forEach((p) => p.classList.toggle('active', p.id === `tab-${id}`));
      if (onChange) onChange(id);
    }
    buttons.forEach((b, i) => {
      b.setAttribute('role', 'tab');
      if (!b.getAttribute('type')) b.setAttribute('type', 'button');
      b.addEventListener('click', () => activate(b.dataset.tab));
      b.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
        const next = buttons[(i + (e.key === 'ArrowRight' ? 1 : buttons.length - 1)) % buttons.length];
        next.focus();
        activate(next.dataset.tab);
      });
    });
    return activate;
  }

  function initials(name) {
    const parts = String(name || '?').replace(/@.*/, '').split(/[\s._-]+/).filter(Boolean);
    return ((parts[0] || '?')[0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
  }

  // ---------- header ----------
  let userSlot = null;

  function toolIconId(id) { return `t-${id}`; }

  function renderHeader(slot, tool) {
    const reg = R();
    const themeUse = document.createElementNS(SVG_NS, 'use');
    themeUse.setAttribute('href', `${ICONS}#${currentTheme() === 'dark' ? 'i-sun' : 'i-moon'}`);
    themeUse.setAttribute('data-tk-theme-icon', '');
    const themeSvg = document.createElementNS(SVG_NS, 'svg');
    themeSvg.setAttribute('class', 'icon');
    themeSvg.setAttribute('aria-hidden', 'true');
    themeSvg.appendChild(themeUse);
    const themeBtn = h('button.btn-icon', { type: 'button', title: 'Switch light or dark mode', 'aria-label': 'Switch light or dark mode', on: { click: () => setTheme(currentTheme() === 'dark' ? 'light' : 'dark') } }, themeSvg);

    const menu = renderMenu(tool);
    userSlot = h('div.tk-user', { hidden: true });

    const left = [
      h('a.tk-brand', { href: ROOT + 'index.html', title: `${reg.TOOLKIT.name} home` }, [
        h('img', { src: SHARED + 'img/jsolve-mark.png', alt: '' }),
        h('span.tk-brand-text', { text: reg.TOOLKIT.name }),
      ]),
    ];
    if (tool) {
      left.push(h('span.tk-crumb-sep', { 'aria-hidden': 'true', text: '/' }));
      left.push(h('div.tk-tool', null, [
        h('span.tk-tool-icon', null, icon(toolIconId(tool.id))),
        h('div.tk-tool-text', null, [
          h('h1.tk-tool-name', { text: tool.id }),
          h('p.tk-tool-tagline', { text: tool.tagline }),
        ]),
      ]));
    } else {
      left[0].classList.add('home');
    }

    const header = h('header.tk-header', null, h('div.tk-header-inner', null, [
      ...left,
      h('div.tk-header-actions', null, [menu, themeBtn, userSlot]),
    ]));
    slot.replaceWith(header);
  }

  function renderMenu(current) {
    const reg = R();
    const panel = h('div.tk-menu-panel', { hidden: true, role: 'menu' });
    for (const g of reg.GROUPS) {
      panel.appendChild(h('div.tk-menu-group', { text: g.label }));
      panel.appendChild(h('div.tk-menu-grid', null, reg.TOOLS.filter((t) => t.group === g.id).map((t) => h('a.tk-menu-item', {
        href: `${ROOT}${t.id}/index.html`, role: 'menuitem', 'aria-current': current && current.id === t.id ? 'page' : null,
      }, [
        h('span.tk-mi-icon', null, icon(toolIconId(t.id))),
        h('span', null, [h('span.tk-mi-name', { text: t.id }), h('span.tk-mi-desc', { text: t.short })]),
      ]))));
    }
    panel.appendChild(h('div.tk-menu-foot', null, [
      h('a', { href: ROOT + 'index.html' }, [icon('i-home', 'sm'), 'All tools']),
      h('a', { href: reg.TOOLKIT.softwareUrl, target: '_blank', rel: 'noopener' }, [icon('i-external', 'sm'), 'More from JSolve']),
    ]));
    const btn = h('button.btn', { type: 'button', 'aria-haspopup': 'true', 'aria-expanded': 'false' }, [icon('i-grid'), h('span.tk-menu-label', { text: 'Tools' }), icon('i-chevron-down', 'sm')]);
    const wrap = h('div.tk-menu', null, [btn, panel]);
    const close = () => { panel.hidden = true; btn.setAttribute('aria-expanded', 'false'); };
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const open = panel.hidden;
      panel.hidden = !open;
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) { const first = panel.querySelector('[aria-current="page"]') || panel.querySelector('a'); if (first) first.focus(); }
    });
    document.addEventListener('click', (e) => { if (!wrap.contains(e.target)) close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !panel.hidden) { close(); btn.focus(); } });
    return wrap;
  }

  // setUser({ name, upn, tenant, onSignOut }) or setUser(null). Shows who is signed in
  // and in which tenant, since tools can hold sessions for more than one.
  function setUser(user) {
    if (!userSlot) return;
    userSlot.replaceChildren();
    if (!user) { userSlot.hidden = true; return; }
    const label = user.upn || user.name || 'Signed in';
    const title = [user.name, user.upn, user.tenant ? `Tenant: ${user.tenant}` : null].filter(Boolean).join('\n');
    userSlot.append(
      h('span.tk-avatar', { text: initials(user.name || user.upn), 'aria-hidden': 'true' }),
      h('span.tk-user-name', { text: label, title }),
    );
    if (user.onSignOut) {
      userSlot.append(h('button.btn-icon.borderless', { type: 'button', title: 'Sign out', 'aria-label': 'Sign out', on: { click: user.onSignOut } }, icon('i-logout')));
    }
    userSlot.hidden = false;
  }

  // ---------- related + footer ----------
  function productCard(item, compact) {
    const reg = R();
    const isTool = reg.TOOLS.includes(item);
    const href = isTool ? `${ROOT}${item.id}/index.html` : item.url;
    const ext = !isTool;
    return h('a.tool-card', { class: compact ? 'compact' : '', href, target: ext ? '_blank' : null, rel: ext ? 'noopener' : null }, [
      h('div.tc-head', null, [
        h('span.tc-icon', null, icon(toolIconId(item.id))),
        h('h3', null, [item.id, ext ? h('span.pill.ext', { text: item.kind }) : null]),
      ]),
      h('p', { text: compact && item.short ? item.short : item.description }),
      h('div.tc-foot', null, [h('span', { text: ext ? item.host : 'Open tool' }), icon(ext ? 'i-external' : 'i-arrow-right')]),
    ]);
  }

  function renderRelated(slot, tool) {
    const reg = R();
    const items = (tool && tool.related ? tool.related : []).map(reg.byId).filter(Boolean);
    if (!items.length) { slot.remove(); return; }
    slot.replaceWith(h('section.tk-related', { 'aria-label': 'Related tools' }, [
      h('h2', { text: 'Related tools' }),
      h('div.tk-related-grid', null, items.map((t) => productCard(t, true))),
    ]));
  }

  function renderFooter(slot, tool) {
    const reg = R();
    const k = reg.TOOLKIT;
    const links = [
      h('a', { href: ROOT + 'index.html' }, [icon('i-home', 'sm'), 'All tools']),
      h('a', { href: tool ? `${k.repoUrl}/tree/main/${tool.id}` : k.repoUrl, target: '_blank', rel: 'noopener' }, [icon('i-github', 'sm'), 'Source']),
      h('a', { href: k.commercialUrl, target: '_blank', rel: 'noopener' }, 'Commercial use'),
      h('a', { href: k.privacyUrl, target: '_blank', rel: 'noopener' }, 'Privacy'),
      h('a', { href: k.ownerUrl, target: '_blank', rel: 'noopener' }, 'jsolve.nl'),
    ];
    slot.replaceWith(h('footer.tk-footer', null, h('div.tk-footer-inner', null, [
      h('div.tk-footer-brand', null, [
        h('img', { src: SHARED + 'img/jsolve-mark.png', alt: '' }),
        h('div', null, [
          h('p', null, [h('strong', { text: `${k.name}` }), ` by ${k.owner}, runs in your browser: your tenant data goes from Microsoft to you${tool && tool.privacyNote ? '.' : ' and nowhere else.'}`]),
          tool && tool.privacyNote ? h('p', { text: tool.privacyNote }) : null,
          h('p', { text: `© ${new Date().getFullYear()} ${k.owner} All rights reserved.` }),
        ]),
      ]),
      h('nav.tk-footer-links', { 'aria-label': 'Toolkit links' }, links),
      h('p.tk-footer-note', { text: 'Provided as-is, use at your own risk. Try changes in a test tenant first and read what a tool is about to do before you confirm it. Free for non-commercial use; commercial use is covered by the commercial use terms.' }),
    ])));
  }

  // Standard "how this tool treats your data" block for About tabs.
  function renderAbout(slot, tool) {
    const block = (ico, title, text) => h('div', null, [icon(ico), h('div', null, [h('strong', { text: title }), text])]);
    const blocks = [
      block('i-lock', 'Runs in your browser', 'No backend. Calls go straight from this page to Microsoft as you, with your own permissions.'),
      tool && tool.readOnly
        ? block('i-eye', 'Read-only', 'This tool only reads. It asks for read permissions and never changes anything in your tenant.')
        : block('i-shield', 'Delegated access only', 'Write permissions are asked for when you first make a change, not before. Nothing changes without a preview and your confirmation.'),
      block('i-history', 'Session-only sign-in', 'Tokens live in this browser tab and are gone when you close it. Nothing is stored on a server.'),
    ];
    if (tool && tool.privacyNote) blocks.push(block('i-info', 'Outside Microsoft', tool.privacyNote));
    slot.replaceWith(h('div.tk-safety', null, blocks));
  }

  function renderDebug(slot) {
    const panel = h('pre#logPanel.log-panel', { 'aria-live': 'off' });
    const verbose = h('input', { type: 'checkbox', id: 'verboseLog' });
    verbose.addEventListener('change', () => window.Log && Log.setVerbose(verbose.checked));
    const copyBtn = h('button.btn.btn-small', { type: 'button', on: { click: async () => {
      try { await navigator.clipboard.writeText(window.Log ? Log.text() : panel.textContent); toast('Log copied', 'ok'); }
      catch { toast('Copy failed, select the text instead', 'warn'); }
    } } }, [icon('i-copy'), 'Copy']);
    const dlBtn = h('button.btn.btn-small', { type: 'button', on: { click: () => window.Log && Log.download() } }, [icon('i-download'), 'Download']);
    const clearBtn = h('button.btn.btn-small', { type: 'button', on: { click: () => window.Log && Log.clear() } }, [icon('i-trash'), 'Clear']);
    slot.replaceWith(h('div.card', null, [
      h('div.card-head', null, [
        h('div', null, [h('h2', { text: 'Debug log' }), h('p.muted.small', { text: 'Every call this tool makes, with timings and errors. Tokens are never logged.' })]),
        h('div.row.tight', null, [h('label.checkbox', null, [verbose, 'Verbose']), copyBtn, dlBtn, clearBtn]),
      ]),
      panel,
    ]));
    if (window.Log && Log.attach) Log.attach(panel);
  }

  function currentTool() {
    const id = document.body && document.body.dataset.tool;
    return id && R() ? R().byId(id) : null;
  }

  function renderShell() {
    if (!R()) { console.error('registry.js must load before toolkit.js renders'); return; }
    const tool = currentTool();
    document.querySelectorAll('[data-tk="header"]').forEach((s) => renderHeader(s, tool));
    document.querySelectorAll('[data-tk="related"]').forEach((s) => renderRelated(s, tool));
    document.querySelectorAll('[data-tk="footer"]').forEach((s) => renderFooter(s, tool));
    document.querySelectorAll('[data-tk="about"]').forEach((s) => renderAbout(s, tool));
    document.querySelectorAll('[data-tk="debug"]').forEach((s) => renderDebug(s));
    hydrateIcons(document);
    if (tool && !document.title) document.title = `${tool.id} - ${R().TOOLKIT.name}`;
    document.dispatchEvent(new CustomEvent('toolkit:ready', { detail: { tool } }));
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', renderShell, { once: true });
  else renderShell();

  window.Toolkit = {
    esc, escAttr: esc, safeUrl, h, icon, iconHtml, hydrateIcons,
    toast, modal, confirm, busy, tabs, setUser, setTheme, initials,
    productCard, paths: { root: ROOT, shared: SHARED, icons: ICONS },
  };
})();
