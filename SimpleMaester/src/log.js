// Lightweight log panel writer. We mirror everything to console too so devtools
// users get one cohesive view.
(() => {
  const panel = () => document.getElementById('logPanel');
  const verbose = () => !!document.getElementById('verboseLog')?.checked;
  const _line = (level, args) => {
    const text = args.map(a => (a instanceof Error ? a.message : (typeof a === 'object' ? JSON.stringify(a) : String(a)))).join(' ');
    const stamp = new Date().toISOString().substring(11, 19);
    const row = `[${stamp}] ${level.padEnd(5)} ${text}`;
    const p = panel();
    if (p) {
      p.textContent += row + '\n';
      p.scrollTop = p.scrollHeight;
    }
    if (level === 'ERR ') console.error(row);
    else if (level === 'WARN') console.warn(row);
    else console.log(row);
  };
  window.Log = {
    info:  (...a) => _line('INFO', a),
    warn:  (...a) => _line('WARN', a),
    err:   (...a) => _line('ERR ', a),
    debug: (...a) => { if (verbose()) _line('DBG ', a); },
    clear: () => { const p = panel(); if (p) p.textContent = ''; },
  };
})();
