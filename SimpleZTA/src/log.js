(() => {
  const lines = [];
  const maxLines = 1500;

  function push(level, parts) {
    const text = `[${new Date().toISOString()}] [${level}] ${parts.map(p => typeof p === 'string' ? p : JSON.stringify(p)).join(' ')}`;
    lines.push(text);
    if (lines.length > maxLines) lines.shift();
    const el = document.getElementById('debugLog');
    if (el) {
      el.textContent = lines.join('\n');
      el.scrollTop = el.scrollHeight;
    }
  }

  const api = {
    info: (...p) => push('INFO', p),
    warn: (...p) => push('WARN', p),
    err: (...p) => push('ERROR', p),
    clear: () => {
      lines.length = 0;
      const el = document.getElementById('debugLog');
      if (el) el.textContent = '';
    },
  };

  window.Log = api;
})();
