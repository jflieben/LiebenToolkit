// Markdown for test descriptions. Uses marked 12 with a renderer that shows raw HTML
// as text and only keeps http(s), mailto and relative links.
(() => {
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (c) => ESC[c]);

  // marked has already entity-encoded href; decode the few entities before checking.
  function cleanHref(href) {
    const s = String(href == null ? '' : href).replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").trim();
    if (!s || /[\u0000-\u001f]/.test(s)) return null;
    let u;
    try { u = new URL(s, location.href); } catch { return null; }
    return ['http:', 'https:', 'mailto:'].includes(u.protocol) ? s : null;
  }

  let instance = null;
  function md() {
    if (instance || !window.marked || !window.marked.Marked) return instance;
    instance = new window.marked.Marked({
      gfm: true,
      breaks: false,
      renderer: {
        html(html) { return esc(typeof html === 'object' && html ? html.text : html); },
        link(href, title, text) {
          const url = cleanHref(href);
          if (!url) return text;
          const ext = /^https?:/i.test(url);
          return `<a href="${esc(url)}"${title ? ` title="${esc(title)}"` : ''}${ext ? ' target="_blank" rel="noopener noreferrer"' : ''}>${text}</a>`;
        },
        image(href, title, text) {
          const url = cleanHref(href);
          if (!url || /^mailto:/i.test(url)) return esc(text);
          return `<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(text || url)}</a>`;
        },
      },
    });
    return instance;
  }

  // Returns trusted HTML. Without marked, the text is escaped with line breaks kept.
  function render(text) {
    const src = String(text == null ? '' : text);
    const m = md();
    if (!m) return src ? `<p>${esc(src).replace(/\n{2,}/g, '</p><p>').replace(/\n/g, '<br>')}</p>` : '';
    return m.parse(src);
  }

  window.Markdown = { render, cleanHref };
})();
