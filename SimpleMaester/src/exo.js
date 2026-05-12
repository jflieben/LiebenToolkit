// Exchange Online admin REST client.
//
// Strategy:
//   1) Try a DIRECT browser call to https://outlook.office365.com/adminapi/...
//      with the user's Exchange access token. If Microsoft has enabled CORS
//      for that endpoint (or we're loaded from an allowed origin), this is
//      the cleanest path - no infrastructure required.
//   2) If the browser blocks the response with a CORS error, transparently
//      fall back to a same-origin proxy at /api/exo (only available in our
//      local PowerShell dev server). For static-hosted deployments without
//      a proxy, surface a clean Skipped result.
//
// The token is acquired in the browser via delegated PKCE for the user.
(() => {
  const TRANSIENT_CODES = new Set([
    'CmdletProxyNotAvailableFailure', 'ServerBusyException',
    'TransientException', 'BackendCommunicationException',
  ]);
  const RESPONSE_CACHE = new Map();
  function _key(org, cmdlet, params) { return org + '|' + cmdlet + '|' + JSON.stringify(params || {}); }
  function clearCache() { RESPONSE_CACHE.clear(); _transport = null; }

  // null = unknown, 'direct' = browser CORS works, 'proxy' = use /api/exo,
  // 'none' = neither available
  let _transport = null;

  async function _detectTransport(token, organization, upn) {
    if (_transport) return _transport;
    // Try direct first. CORS preflight will fail fast if the endpoint doesn't
    // allow the SPA origin.
    try {
      const r = await _callDirect(token, organization, upn, 'Get-OrganizationConfig', {});
      // Any HTTP response (even 401/403) means CORS allowed the request through.
      if (r.status > 0) {
        _transport = 'direct';
        Log.info('Exo: using direct browser call (CORS allowed by outlook.office365.com).');
        // Drain the probe response so it doesn't leak.
        try { await r.response.text(); } catch { /* ignore */ }
        return _transport;
      }
    } catch (e) {
      Log.warn(`Exo: direct call blocked (${e.message}), trying same-origin proxy /api/exo...`);
    }
    // Fall back to proxy.
    try {
      const res = await fetch('/api/exo', { method: 'OPTIONS' });
      if (res.status === 204 || res.status === 200 || res.status === 405) {
        _transport = 'proxy';
        Log.info('Exo: using same-origin proxy at /api/exo.');
        return _transport;
      }
    } catch { /* ignore */ }
    _transport = 'none';
    return _transport;
  }

  async function _callDirect(token, organization, upn, cmdletName, parameters) {
    const url = `https://outlook.office365.com/adminapi/beta/${encodeURIComponent(organization)}/InvokeCommand`;
    const body = JSON.stringify({
      CmdletInput: {
        CmdletName: cmdletName,
        Parameters: parameters || {},
      },
    });
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json; charset=utf-8',
      'Accept': 'application/json',
      'X-AnchorMailbox': `UPN:${upn}`,
      'X-ResponseFormat': 'json',
      'Prefer': 'odata.maxpagesize=1000',
    };
    const t0 = performance.now();
    const res = await fetch(url, { method: 'POST', headers, body, mode: 'cors' });
    const dur = Math.round(performance.now() - t0);
    return { status: res.status, response: res, durationMs: dur };
  }

  async function _callProxy(token, organization, upn, cmdletName, parameters) {
    const proxyBody = {
      token, organization, upn, cmdlet: cmdletName, parameters: parameters || {},
    };
    const t0 = performance.now();
    const res = await fetch('/api/exo', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(proxyBody),
    });
    const dur = Math.round(performance.now() - t0);
    return { status: res.status, response: res, durationMs: dur };
  }

  async function invoke(cmdletName, parameters, organization, opts = {}) {
    if (!organization) throw new Error(`Exo.invoke(${cmdletName}): organization (tenant domain) is required`);

    const ck = _key(organization, cmdletName, parameters);
    if (RESPONSE_CACHE.has(ck) && !opts.bypassCache) {
      const cached = RESPONSE_CACHE.get(ck);
      const sh = ApiRecorder.shorten(cached);
      ApiRecorder.record({ method: 'EXO', url: cmdletName, status: 200, durationMs: 0, fromCache: true, response: sh.response, truncated: sh.truncated });
      return cached;
    }

    const tok = await Auth.getExchangeToken();
    const upn = Auth.getUserPrincipalName();
    if (!upn) throw new Error('Exo.invoke: no signed-in user (UPN)');

    const transport = await _detectTransport(tok, organization, upn);
    if (transport === 'none') {
      const msg = `Exchange admin API needs either CORS allowed on outlook.office365.com (not the case for browser SPAs by default) or a same-origin proxy at /api/exo (not deployed here).`;
      ApiRecorder.record({ method: 'EXO', url: cmdletName, status: 0, durationMs: 0, fromCache: false, response: msg, truncated: false });
      const err = new Error(msg);
      err.code = 'CORS_BLOCKED';
      err.status = 0;
      throw err;
    }

    let attempt = 0;
    while (true) {
      let result;
      try {
        result = transport === 'direct'
          ? await _callDirect(tok, organization, upn, cmdletName, parameters)
          : await _callProxy(tok, organization, upn, cmdletName, parameters);
      } catch (netErr) {
        const msg = `Exo network error (${cmdletName}): ${netErr.message}`;
        ApiRecorder.record({ method: 'EXO', url: cmdletName, status: 0, durationMs: 0, fromCache: false, response: msg, truncated: false });
        const err = new Error(msg);
        err.code = 'CORS_BLOCKED';
        err.status = 0;
        throw err;
      }

      const { response: res, durationMs: dur } = result;

      if (res.status === 429 || res.status === 503 || res.status === 504) {
        attempt++;
        if (attempt > 5) {
          ApiRecorder.record({ method: 'EXO', url: cmdletName, status: res.status, durationMs: dur, fromCache: false, response: 'throttled, retries exhausted', truncated: false });
          throw new Error(`Exchange ${res.status} for ${cmdletName}`);
        }
        const wait = parseInt(res.headers.get('Retry-After') || '0', 10) * 1000 || (Math.pow(2, attempt) * 1000);
        Log.warn(`Exo ${res.status} for ${cmdletName}, sleeping ${wait}ms`);
        await sleep(wait);
        continue;
      }

      const text = await res.text().catch(() => '');
      let payload = null;
      if (text) { try { payload = JSON.parse(text); } catch { /* ignore */ } }

      if (!res.ok) {
        const code = payload?.error?.details?.[0]?.code || payload?.error?.code;
        if (code && TRANSIENT_CODES.has(code)) {
          attempt++;
          if (attempt <= 4) {
            await sleep(Math.pow(2, attempt) * 1000);
            continue;
          }
        }
        const sh = ApiRecorder.shorten(text);
        ApiRecorder.record({ method: 'EXO', url: cmdletName, status: res.status, durationMs: dur, fromCache: false, response: sh.response, truncated: sh.truncated });
        const err = new Error(`Exchange ${res.status} for ${cmdletName}: ${text.substring(0, 400)}`);
        err.status = res.status;
        err.body = text;
        err.code = code;
        throw err;
      }

      const value = Array.isArray(payload?.value) ? payload.value : (payload ? [payload] : []);
      const sh = ApiRecorder.shorten(value);
      ApiRecorder.record({ method: 'EXO', url: cmdletName, status: res.status, durationMs: dur, fromCache: false, response: sh.response, truncated: sh.truncated });
      RESPONSE_CACHE.set(ck, value);
      return value;
    }
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  window.Exo = { invoke, clearCache };
})();
