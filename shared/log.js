// Shared debug log. Writes to the panel attached via Log.attach() (the shell's debug
// card does that) or #logPanel. Anything that looks like a token is masked first.
(() => {
  const MAX = 5000;
  const state = { verbose: false, entries: [], panel: null };

  // Bearer headers, JWTs and token-ish query parameters.
  const SECRET_PATTERNS = [
    [/Bearer\s+[A-Za-z0-9\-._~+/]+=*/gi, 'Bearer [redacted]'],
    [/eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{4,}/g, '[jwt redacted]'],
    [/((?:access_token|refresh_token|id_token|code|client_secret|password)=)[^&\s"']+/gi, '$1[redacted]'],
    [/("(?:access_token|refresh_token|id_token|accessToken|refreshToken|idToken|secret|password)"\s*:\s*")[^"]*"/gi, '$1[redacted]"'],
  ];
  function redact(text) {
    let out = text;
    for (const [re, rep] of SECRET_PATTERNS) out = out.replace(re, rep);
    return out;
  }

  function format(level, args) {
    const ts = new Date().toISOString().substring(11, 23);
    const text = args.map((a) => {
      if (a instanceof Error) return a.stack || a.message;
      if (a && typeof a === 'object') { try { return JSON.stringify(a); } catch { return String(a); } }
      return String(a);
    }).join(' ');
    return { ts, level, text: redact(text) };
  }

  function panel() {
    if (state.panel && state.panel.isConnected) return state.panel;
    state.panel = document.getElementById('logPanel');
    return state.panel;
  }

  function line(entry) {
    const div = document.createElement('div');
    div.className = `log-${entry.level}`;
    div.textContent = `[${entry.ts}] ${entry.level.toUpperCase().padEnd(4)} ${entry.text}`;
    return div;
  }

  function append(entry) {
    state.entries.push(entry);
    if (state.entries.length > MAX) state.entries.shift();
    const p = panel();
    if (!p || (entry.level === 'dbg' && !state.verbose)) return;
    const stick = p.scrollTop + p.clientHeight >= p.scrollHeight - 30;
    p.appendChild(line(entry));
    while (p.childElementCount > MAX) p.firstChild.remove();
    if (stick) p.scrollTop = p.scrollHeight;
  }

  function rerender() {
    const p = panel();
    if (!p) return;
    p.replaceChildren(...state.entries.filter((e) => state.verbose || e.level !== 'dbg').map(line));
    p.scrollTop = p.scrollHeight;
  }

  function text() {
    return state.entries.map((e) => `[${e.ts}] ${e.level.toUpperCase().padEnd(4)} ${e.text}`).join('\n');
  }

  window.Log = {
    dbg: (...a) => append(format('dbg', a)),
    info: (...a) => { append(format('info', a)); console.log(...a); },
    warn: (...a) => { append(format('warn', a)); console.warn(...a); },
    err: (...a) => { append(format('err', a)); console.error(...a); },
    error: (...a) => window.Log.err(...a),
    setVerbose(v) { state.verbose = !!v; rerender(); },
    isVerbose() { return state.verbose; },
    attach(el) { state.panel = el; rerender(); },
    getEntries() { return state.entries.slice(); },
    text,
    download(name) {
      const tool = (document.body && document.body.dataset.tool) || 'toolkit';
      const blob = new Blob([text()], { type: 'text/plain;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name || `${tool}-log-${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    },
    clear() { state.entries = []; const p = panel(); if (p) p.replaceChildren(); },
    redact,
  };
})();
