(() => {
  const panel = () => document.getElementById('logPanel');
  let verbose = false;

  function now() {
    return new Date().toISOString().replace('T', ' ').replace('Z', '');
  }

  function write(level, args) {
    if (level === 'DBG' && !verbose) return;
    const p = panel();
    const msg = Array.from(args)
      .map((v) => (typeof v === 'string' ? v : safeJson(v)))
      .join(' ');
    if (p) {
      p.textContent += `[${now()}] ${level} ${msg}\n`;
      p.scrollTop = p.scrollHeight;
    }
    const fn = level === 'ERR' ? console.error : level === 'WRN' ? console.warn : console.log;
    fn('[ConditionalAccess]', msg);
  }

  function safeJson(v) {
    try { return JSON.stringify(v); } catch { return String(v); }
  }

  window.Log = {
    dbg(...args) { write('DBG', args); },
    info(...args) { write('INF', args); },
    warn(...args) { write('WRN', args); },
    err(...args) { write('ERR', args); },
    clear() {
      const p = panel();
      if (p) p.textContent = '';
    },
    setVerbose(v) {
      verbose = !!v;
      write('INF', [`Verbose logging ${verbose ? 'enabled' : 'disabled'}.`]);
    },
  };
})();