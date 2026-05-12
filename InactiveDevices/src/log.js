// Centralized logging.
(() => {
  const state = { verbose: false, entries: [] };

  function format(level, args) {
    const ts = new Date().toISOString().substring(11, 23);
    const text = args.map((a) => {
      if (a instanceof Error) return a.stack || a.message;
      if (typeof a === 'object') {
        try { return JSON.stringify(a); } catch { return String(a); }
      }
      return String(a);
    }).join(' ');
    return { ts, level, text };
  }

  function append(entry) {
    state.entries.push(entry);
    if (state.entries.length > 5000) state.entries.shift();
    const panel = document.getElementById('logPanel');
    if (!panel) return;
    if (entry.level === 'dbg' && !state.verbose) return;
    const line = document.createElement('div');
    line.className = `log-${entry.level}`;
    line.textContent = `[${entry.ts}] ${entry.level.toUpperCase().padEnd(4)} ${entry.text}`;
    panel.appendChild(line);
    panel.scrollTop = panel.scrollHeight;
  }

  window.Log = {
    dbg: (...a) => append(format('dbg', a)),
    info: (...a) => { append(format('info', a)); console.log(...a); },
    warn: (...a) => { append(format('warn', a)); console.warn(...a); },
    err: (...a) => { append(format('err', a)); console.error(...a); },
    setVerbose(v) { state.verbose = v; },
    getEntries() { return state.entries.slice(); },
    clear() { state.entries = []; const p = document.getElementById('logPanel'); if (p) p.textContent = ''; },
  };
})();
