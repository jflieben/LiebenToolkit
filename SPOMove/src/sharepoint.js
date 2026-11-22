// SharePoint REST client (per-connection: role = 'source' | 'target').
// All calls use the SharePoint token for the connection's tenant so the same
// helpers work for same-tenant and cross-tenant moves.
(() => {
  const CHUNK_SIZE = 8 * 1024 * 1024;      // 8 MB upload chunks
  const CHUNK_THRESHOLD = 10 * 1024 * 1024; // switch to chunked upload above this

  // Percent-encode a server-relative path for use inside decodedurl='...'.
  // Single quotes are doubled per OData string-literal rules first.
  function spPath(p) { return encodeURIComponent(String(p).replace(/'/g, "''")); }

  function trimUrl(u) { return String(u || '').replace(/\/+$/, ''); }

  // ---------- server-relative path helpers ----------
  function serverRelativeOf(absoluteUrl) {
    try {
      const p = new URL(absoluteUrl).pathname;
      try { return decodeURIComponent(p); } catch { return p; }
    } catch { return absoluteUrl; }
  }
  function hostOf(absoluteUrl) {
    try { return new URL(absoluteUrl).host; } catch { return ''; }
  }
  function folderOf(serverRelPath) {
    const i = serverRelPath.lastIndexOf('/');
    return i <= 0 ? '/' : serverRelPath.substring(0, i);
  }
  function nameOf(serverRelPath) {
    const i = serverRelPath.lastIndexOf('/');
    return i < 0 ? serverRelPath : serverRelPath.substring(i + 1);
  }
  function joinPath(folder, name) { return `${trimUrl(folder)}/${name}`; }

  // ---------- adaptive throttle circuit breaker (shared by ALL requests / jobs) ----------
  // Healthy: every request passes. On throttle (429 / dropped connection) the circuit OPENS
  // and every request parks behind a shared backoff. When the backoff elapses it goes
  // HALF-OPEN: a single *random* request probes. On success the allowance ramps up (more
  // requests are let through) until fully reopened; on failure the backoff grows
  // (multiplicatively) and a fresh random probe is tried next round. Only the current
  // round's probe escalates the backoff, so a herd of simultaneous 429s can't keep
  // re-strengthening it, and randomizing the prober avoids hanging on one bad request.
  const INITIAL_BACKOFF = 5000;
  const MAX_BACKOFF = 120000;
  const HEALTHY_AT = 8;            // recovery allowance at which the circuit fully reopens
  const _tc = { backoffMs: 0, gateUntil: 0, permits: 0, inFlight: 0, epoch: 0 };

  // Rate-limit noisy throttle warnings to one every 5s (with a suppressed count).
  let _lastThrottleLog = 0;
  let _throttleSuppressed = 0;
  function _logThrottle(msg) {
    const now = Date.now();
    if (now - _lastThrottleLog > 5000) {
      Log.warn(msg + (_throttleSuppressed ? ` (+${_throttleSuppressed} similar suppressed)` : ''));
      _lastThrottleLog = now;
      _throttleSuppressed = 0;
    } else {
      _throttleSuppressed++;
    }
  }

  async function _throttleAcquire() {
    while (true) {
      if (_tc.backoffMs === 0) return { recovery: false, epoch: _tc.epoch };
      const now = Date.now();
      if (now < _tc.gateUntil) {
        // Backoff window: wait in randomized chunks so threads wake near expiry in a random
        // order; whoever wakes first grabs the single probe slot.
        const remaining = _tc.gateUntil - now;
        await Concurrency.sleep(Math.min(remaining, 1000) + Math.random() * 300);
        continue;
      }
      if (_tc.permits === 0) _tc.permits = 1;                 // open one probe slot
      if (_tc.inFlight < _tc.permits) { _tc.inFlight++; return { recovery: true, epoch: _tc.epoch }; }
      await Concurrency.sleep(50 + Math.random() * 300);       // contend randomly for the slot
    }
  }

  function _throttleReport(ok, retryAfterSec, slot) {
    if (slot.recovery) _tc.inFlight = Math.max(0, _tc.inFlight - 1);
    const now = Date.now();
    if (ok) {
      // Server responded (2xx, or even a 4xx like 404): not throttling.
      if (_tc.backoffMs !== 0 && slot.recovery && slot.epoch === _tc.epoch) {
        _tc.permits += 1;                                     // ramp the allowance up
        if (_tc.permits >= HEALTHY_AT) {
          _tc.backoffMs = 0; _tc.gateUntil = 0; _tc.permits = 0; _tc.inFlight = 0;
          Log.info('SPO throttling cleared, resuming full speed.');
        }
      }
      return;
    }
    const raMs = retryAfterSec ? retryAfterSec * 1000 : 0;
    if (_tc.backoffMs === 0) {
      // First throttle: open the circuit with a substantial initial backoff.
      _tc.backoffMs = Math.min(MAX_BACKOFF, Math.max(INITIAL_BACKOFF, raMs));
      _tc.gateUntil = now + _tc.backoffMs; _tc.permits = 0; _tc.inFlight = 0; _tc.epoch++;
      _logThrottle(`SPO throttled, pausing all requests, backing off ${Math.round(_tc.backoffMs / 1000)}s`);
    } else if (slot.recovery && slot.epoch === _tc.epoch) {
      // The current round's probe failed: escalate and reset the ramp.
      _tc.backoffMs = Math.min(MAX_BACKOFF, Math.max(_tc.backoffMs * 2, raMs));
      _tc.gateUntil = now + _tc.backoffMs; _tc.permits = 0; _tc.inFlight = 0; _tc.epoch++;
      _logThrottle(`SPO still throttled, backing off ${Math.round(_tc.backoffMs / 1000)}s`);
    } else if (raMs) {
      // Stale / herd 429 while already backing off: honor a longer Retry-After, never escalate.
      _tc.gateUntil = Math.max(_tc.gateUntil, now + raMs);
    }
  }

  // Snapshot of the shared throttle circuit for the UI. throttled while backing off;
  // waitMs counts down to the next retry during the backoff window (0 while probing).
  function getThrottleState() {
    if (_tc.backoffMs === 0) return { throttled: false, phase: 'ok', waitMs: 0 };
    const waitMs = Math.max(0, _tc.gateUntil - Date.now());
    return { throttled: true, phase: waitMs > 0 ? 'backoff' : 'recovering', waitMs, backoffMs: _tc.backoffMs };
  }

  // ---------- generic REST ----------
  // Http.request does the fetch with retries off: the circuit breaker above owns throttling
  // for all SharePoint calls of all jobs. The token is fetched first so an auth error is
  // never mistaken for a dropped connection.
  async function callRest(role, siteUrl, relativePath, { method = 'GET', body, headers = {}, raw = false, digest, attempt = 0, signal } = {}) {
    const token = await Auth.getSpoToken(role, hostOf(siteUrl));
    const slot = await _throttleAcquire();
    const url = trimUrl(siteUrl) + relativePath;
    const h = { Accept: 'application/json;odata=nometadata', ...headers };
    if (digest) h['X-RequestDigest'] = digest;
    const again = () => callRest(role, siteUrl, relativePath, { method, body, headers, raw, digest, attempt: attempt + 1, signal });

    let res;
    try {
      res = await Http.request(url, { method, headers: h, body: body == null ? undefined : body, token, retries: 0, raw: true, signal, label: `SPO[${role}]` });
    } catch (e) {
      if (/^Cancelled$/.test(e.message || '')) {
        if (slot.recovery) _tc.inFlight = Math.max(0, _tc.inFlight - 1);
        throw e;
      }
      // A dropped connection ("Failed to fetch") under load is almost always throttling.
      _throttleReport(false, 0, slot);
      if (attempt >= 10) throw errWith(`SPO throttled (connection dropped) after retries: ${url}`, 0);
      return again();
    }

    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      _throttleReport(false, ra, slot);
      if (attempt >= 10) throw errWith(`SPO ${res.status} after retries: ${url}`, res.status);
      return again();
    }

    // Any answer, even a 4xx, means the server is reachable and not throttling.
    _throttleReport(true, 0, slot);
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Http.HttpError(res.status, `SPO ${Http.errorMessage(res.status, res.statusText, txt)}`, txt, url);
    }
    if (raw) return res;
    if (res.status === 204) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    return res.text();
  }

  function errWith(msg, status) { const e = new Error(msg); e.status = status; return e; }

  const _digestCache = new Map(); // key = role|siteUrl -> { value, exp }
  const _digestInflight = new Map(); // key -> Promise (coalesce concurrent cold fetches)
  async function getFormDigest(role, siteUrl) {
    const key = `${role}|${trimUrl(siteUrl)}`;
    const cached = _digestCache.get(key);
    if (cached && cached.exp > Date.now()) return cached.value;
    // A burst of parallel writes starting cold would each POST contextinfo; share one.
    if (_digestInflight.has(key)) return _digestInflight.get(key);
    const p = (async () => {
      try {
        const data = await callRest(role, siteUrl, '/_api/contextinfo', {
          method: 'POST', headers: { Accept: 'application/json;odata=verbose' },
        });
        const info = data?.d?.GetContextWebInformation || data?.GetContextWebInformation || data;
        const value = info?.FormDigestValue || info?.FormDigestValue;
        const timeout = (info?.FormDigestTimeoutSeconds || 1800) * 1000;
        if (!value) throw new Error('Could not obtain form digest');
        _digestCache.set(key, { value, exp: Date.now() + timeout - 60000 });
        return value;
      } finally {
        _digestInflight.delete(key);
      }
    })();
    _digestInflight.set(key, p);
    return p;
  }

  // ---------- connection validation ----------
  async function validateConnection(role, siteUrl) {
    const conn = Auth.getConnection(role);
    if (!conn?.account) return { ok: false, error: 'Not signed in' };
    if (!conn.tenantHost) return { ok: false, error: 'SharePoint host not detected for this tenant' };
    const testUrl = siteUrl || `https://${conn.tenantHost}`;
    try {
      const web = await callRest(role, testUrl, '/_api/web?$select=Title,Url');
      return { ok: true, title: web.Title, url: web.Url };
    } catch (e) {
      return { ok: false, status: e.status, error: e.message };
    }
  }

  // ---------- site & library discovery ----------
  async function getWeb(role, siteUrl) {
    return callRest(role, siteUrl, '/_api/web?$select=Title,Url,ServerRelativeUrl');
  }

  async function getDocumentLibraries(role, siteUrl) {
    // Fetch all lists and filter client-side: a server-side $filter=BaseType eq 1 has been
    // seen to return nothing on some target tenants. BaseType 1 = document library. A
    // OneDrive personal site's main "Documents" library reports IsSystemList=true, so we do
    // NOT exclude on IsSystemList; instead drop catalogs, hidden libraries, and a few
    // well-known system libraries by their locale-independent RootFolder URL leaf.
    const path = "/_api/web/lists?$select=Title,BaseType,ItemCount,Hidden,IsCatalog,RootFolder/ServerRelativeUrl,RootFolder/Name&$expand=RootFolder&$top=500";
    const raw = await getAllPaged(role, siteUrl, path);
    const SYSTEM = new Set(['siteassets', 'style library', 'formservertemplates', 'preservationholdlibrary']);
    const isCatalogPath = (u) => /\/_catalogs\//i.test(u || '');
    const toLib = (l) => ({ title: l.Title, itemCount: l.ItemCount, serverRelativeUrl: l.RootFolder.ServerRelativeUrl });
    const pick = (includeHidden) => raw.filter(l => {
      if (l.BaseType !== 1) return false;
      const url = l.RootFolder?.ServerRelativeUrl;
      if (!url || l.IsCatalog || isCatalogPath(url)) return false;
      if (l.Hidden && !includeHidden) return false;
      const leaf = (l.RootFolder?.Name || url.split('/').pop() || '').toLowerCase();
      return !SYSTEM.has(leaf);
    }).map(toLib);

    let libs = pick(false);
    // Some OneDrive/site tenants mark the default document library Hidden; retry with hidden.
    if (!libs.length) libs = pick(true);
    // Last resort: the lists endpoint occasionally omits the default library entirely; fetch it
    // directly by its well-known URL so the user isn't left with "no libraries".
    if (!libs.length) {
      const def = await probeDefaultLibrary(role, siteUrl).catch(() => null);
      if (def) libs = [def];
    }
    Log.dbg(`getDocumentLibraries(${siteUrl}): ${raw.length} raw lists -> ${libs.length} libraries`);
    return libs;
  }

  // Directly fetch a site's default document library by its root-folder URL. OneDrive personal
  // sites use "Documents"; classic/team sites use "Shared Documents". Returns null if neither
  // resolves.
  async function probeDefaultLibrary(role, siteUrl) {
    let webRel;
    try {
      const web = await getWeb(role, siteUrl);
      webRel = trimUrl(web.ServerRelativeUrl || serverRelativeOf(web.Url));
    } catch { webRel = trimUrl(serverRelativeOf(siteUrl)); }
    for (const leaf of ['Documents', 'Shared Documents']) {
      const listUrl = `${webRel}/${leaf}`;
      try {
        const data = await callRest(role, siteUrl, `/_api/web/GetList(@a)?@a='${spPath(listUrl)}'&$select=Title,BaseType,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
        if (data && data.BaseType === 1 && data.RootFolder?.ServerRelativeUrl) {
          return { title: data.Title, itemCount: data.ItemCount, serverRelativeUrl: data.RootFolder.ServerRelativeUrl };
        }
      } catch { /* try the next well-known leaf */ }
    }
    return null;
  }

  // Follow SharePoint's odata.nextLink to page through a collection completely.
  async function getAllPaged(role, siteUrl, relativePath) {
    const out = [];
    let path = relativePath;
    let guard = 0;
    while (path && guard++ < 100000) {
      const data = await callRest(role, siteUrl, path);
      for (const r of (data.value || [])) out.push(r);
      const next = data['odata.nextLink'] || data['@odata.nextLink'];
      if (!next) break;
      path = next.startsWith('http') ? next.substring(trimUrl(siteUrl).length) : (next.startsWith('/') ? next : `/${next}`);
    }
    return out;
  }

  // List immediate children (folders + files) of a folder. Folders and files are
  // paged separately so folders holding more than the list-view page size are still
  // enumerated fully (a single $expand caps the collections and silently truncates).
  async function getFolderChildren(role, siteUrl, serverRelFolder) {
    const enc = spPath(serverRelFolder);
    const base = `/_api/web/GetFolderByServerRelativePath(decodedurl='${enc}')`;
    const [folderRows, fileRows] = await Promise.all([
      getAllPaged(role, siteUrl, `${base}/Folders?$select=Name,ServerRelativeUrl,ItemCount&$top=5000`),
      getAllPaged(role, siteUrl, `${base}/Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified,UIVersionLabel&$top=5000`),
    ]);
    const folders = folderRows
      .filter(f => f.Name !== 'Forms' && !f.Name.startsWith('_'))
      .map(f => ({ name: f.Name, serverRelativeUrl: f.ServerRelativeUrl, itemCount: f.ItemCount }));
    const files = fileRows.map(f => ({
      name: f.Name,
      serverRelativeUrl: f.ServerRelativeUrl,
      size: Number(f.Length || 0),
      created: f.TimeCreated,
      modified: f.TimeLastModified,
      versionLabel: f.UIVersionLabel,
    }));
    return { folders, files };
  }

  // Recursively enumerate every file under a root folder with a pool of parallel folder
  // workers (folder navigation is not subject to the list-view threshold). Resolves with
  // { files, failedFolders: [{ folder, error }] }: a folder that cannot be listed is reported,
  // never silently treated as empty, because a sync would read that as "deleted from source".
  async function enumerateFiles(role, siteUrl, rootFolder, { onProgress, cancelToken, recursive = true, concurrency = 6 } = {}) {
    const files = [];
    const failedFolders = [];
    const queue = [rootFolder];
    let active = 0;
    return new Promise((resolve, reject) => {
      let settled = false;
      const finish = (err) => { if (!settled) { settled = true; err ? reject(err) : resolve({ files, failedFolders }); } };
      function pump() {
        if (settled) return;
        if (cancelToken?.cancelled) return finish(new Error('Cancelled'));
        if (!queue.length && active === 0) return finish();
        while (active < concurrency && queue.length) {
          const folder = queue.shift();
          active++;
          getFolderChildren(role, siteUrl, folder).then(children => {
            for (const f of children.files) {
              files.push({ ...f, relativeTo: rootFolder, subPath: f.serverRelativeUrl.substring(rootFolder.length).replace(/^\/+/, '') });
            }
            if (recursive) for (const d of children.folders) queue.push(d.serverRelativeUrl);
          }).catch(e => {
            if (/^Cancelled$/.test(e.message || '')) return;
            failedFolders.push({ folder, error: e.message, status: e.status });
            Log.warn(`Could not list folder ${folder}: ${e.message}`);
          }).finally(() => {
            active--;
            if (onProgress) onProgress(files.length, queue.length);
            pump();
          });
        }
      }
      pump();
    });
  }

  // ---------- folder creation ----------
  // Create a SINGLE folder (its parent must already exist). Tolerant of a folder that
  // already exists (resume / pre-existing), verified via a follow-up existence check.
  async function createFolder(role, siteUrl, serverRelFolder) {
    const digest = await getFormDigest(role, siteUrl);
    const enc = spPath(serverRelFolder);
    try {
      await callRest(role, siteUrl, `/_api/web/folders/AddUsingPath(decodedurl='${enc}')`, { method: 'POST', digest });
    } catch (e) {
      if (await folderExists(role, siteUrl, serverRelFolder)) return;
      throw e;
    }
  }
  async function folderExists(role, siteUrl, serverRelFolder) {
    try {
      const data = await callRest(role, siteUrl, `/_api/web/GetFolderByServerRelativePath(decodedurl='${spPath(serverRelFolder)}')?$select=Exists`);
      return !data || data.Exists !== false;
    } catch { return false; }
  }

  // ---------- download / upload ----------
  async function downloadFile(role, siteUrl, serverRelPath) {
    const res = await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${spPath(serverRelPath)}')/$value`, { raw: true });
    return res.blob();
  }

  async function uploadFileSmall(role, siteUrl, folderServerRel, name, blob, overwrite = true) {
    const digest = await getFormDigest(role, siteUrl);
    const enc = spPath(folderServerRel);
    const path = `/_api/web/GetFolderByServerRelativePath(decodedurl='${enc}')/Files/AddUsingPath(decodedurl='${spPath(name)}',overwrite=${overwrite ? 'true' : 'false'})`;
    const buf = await blob.arrayBuffer();
    await callRest(role, siteUrl, path, {
      method: 'POST', digest, body: buf,
      headers: { 'Content-Type': 'application/octet-stream' },
    });
  }

  async function uploadFileChunked(role, siteUrl, folderServerRel, name, blob, { overwrite = true, onProgress, cancelToken } = {}) {
    const digest = await getFormDigest(role, siteUrl);
    const fullPath = joinPath(folderServerRel, name);
    // Create/overwrite an empty file first.
    const createPath = `/_api/web/GetFolderByServerRelativePath(decodedurl='${spPath(folderServerRel)}')/Files/AddUsingPath(decodedurl='${spPath(name)}',overwrite=${overwrite ? 'true' : 'false'})`;
    await callRest(role, siteUrl, createPath, {
      method: 'POST', digest, body: new ArrayBuffer(0),
      headers: { 'Content-Type': 'application/octet-stream' },
    });

    const uploadId = (crypto.randomUUID ? crypto.randomUUID() : fallbackGuid());
    const total = blob.size;
    let offset = 0;
    let first = true;
    while (offset < total) {
      if (cancelToken?.cancelled) throw new Error('Cancelled');
      const end = Math.min(offset + CHUNK_SIZE, total);
      const isLast = end >= total;
      const chunk = await blob.slice(offset, end).arrayBuffer();
      const encFile = spPath(fullPath);
      let op;
      if (first) op = `StartUpload(uploadId=guid'${uploadId}')`;
      else if (isLast) op = `FinishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`;
      else op = `ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`;
      const d = await getFormDigest(role, siteUrl);
      await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${encFile}')/${op}`, {
        method: 'POST', digest: d, body: chunk,
        headers: { 'Content-Type': 'application/octet-stream' },
      });
      offset = end;
      first = false;
      if (onProgress) onProgress(offset, total);
    }
    // Zero-length files never enter the loop; the empty create already covered them.
  }

  function fallbackGuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
      const r = Math.random() * 16 | 0; const v = c === 'x' ? r : (r & 0x3 | 0x8); return v.toString(16);
    });
  }

  async function uploadFile(role, siteUrl, folderServerRel, name, blob, opts = {}) {
    if (blob.size > CHUNK_THRESHOLD) return uploadFileChunked(role, siteUrl, folderServerRel, name, blob, opts);
    return uploadFileSmall(role, siteUrl, folderServerRel, name, blob, opts.overwrite !== false);
  }

  // ---------- metadata ----------
  async function getFileMetadata(role, siteUrl, serverRelPath) {
    const enc = spPath(serverRelPath);
    // Read the File's own Author/ModifiedBy users + timestamps. Expanding Author/Editor
    // on ListItemAllFields fails on some libraries ("The $expand query is not valid for
    // field 'Author'"), so use the File-level navigation properties instead.
    const path = `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')` +
      `?$select=TimeCreated,TimeLastModified,Author/Title,Author/Email,Author/LoginName,ModifiedBy/Title,ModifiedBy/Email,ModifiedBy/LoginName&$expand=Author,ModifiedBy`;
    const data = await callRest(role, siteUrl, path);
    return {
      created: data.TimeCreated,
      modified: data.TimeLastModified,
      author: data.Author ? { title: data.Author.Title, email: data.Author.Email, loginName: data.Author.LoginName } : null,
      editor: data.ModifiedBy ? { title: data.ModifiedBy.Title, email: data.ModifiedBy.Email, loginName: data.ModifiedBy.LoginName } : null,
    };
  }

  // Resolve a principal to its site user id. Cached per role|site|logonName because
  // the SAME users repeat across a job's files (and reset-metadata reuses ONE user for
  // every file). Concurrent lookups of the same principal are coalesced into one call.
  const _ensureUserCache = new Map();    // key = role|site|logonName -> { id, loginName }
  const _ensureUserInflight = new Map();
  async function ensureUser(role, siteUrl, logonName) {
    const key = `${role}|${trimUrl(siteUrl)}|${logonName}`;
    const hit = _ensureUserCache.get(key);
    if (hit) return hit;
    if (_ensureUserInflight.has(key)) return _ensureUserInflight.get(key);
    const p = (async () => {
      try {
        const digest = await getFormDigest(role, siteUrl);
        const data = await callRest(role, siteUrl, '/_api/web/ensureuser', {
          method: 'POST', digest,
          headers: { 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify({ logonName }),
        });
        const user = { id: data?.Id ?? null, loginName: data?.LoginName || logonName };
        _ensureUserCache.set(key, user);
        return user;
      } finally {
        _ensureUserInflight.delete(key);
      }
    })();
    _ensureUserInflight.set(key, p);
    return p;
  }

  // Apply Created/Modified and (optionally) Author/Editor to a target file's list
  // item via a MERGE on ListItemAllFields. Dates are sent as ISO 8601, which is
  // culture-invariant; ValidateUpdateListItem instead parses dates with the web's
  // locale and rejects ISO on non-English sites ("valid date within the range of
  // 1-1-1900 ..."). Author/Editor are best-effort: the user must resolve in the
  // target web, else those fields are skipped.
  async function setFileMetadata(role, siteUrl, serverRelPath, meta, { includeAuthorEditor = true } = {}) {
    const fields = {};
    if (meta.created) fields.Created = toSpDate(meta.created);
    if (meta.modified) fields.Modified = toSpDate(meta.modified);

    if (includeAuthorEditor) {
      for (const [idField, principal] of [['AuthorId', meta.author], ['EditorId', meta.editor]]) {
        const claim = principalClaim(principal);
        if (!claim) continue;
        try {
          const user = await ensureUser(role, siteUrl, claim);
          if (user.id != null) fields[idField] = user.id;
        } catch (e) {
          Log.warn(`Could not resolve ${idField} '${claim}' on target: ${e.message}`);
        }
      }
    }
    if (!Object.keys(fields).length) return;

    const digest = await getFormDigest(role, siteUrl);
    const enc = spPath(serverRelPath);
    await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')/ListItemAllFields`, {
      method: 'POST', digest,
      headers: { 'Content-Type': 'application/json;odata=nometadata', 'X-HTTP-Method': 'MERGE', 'If-Match': '*' },
      body: JSON.stringify(fields),
    });
  }

  function principalClaim(principal) {
    if (!principal) return null;
    const email = principal.email;
    const login = principal.loginName || '';
    if (/^i:0/i.test(login)) return login;      // already a claim
    if (email) return `i:0#.f|membership|${email}`;
    return null;
  }

  function toSpDate(v) {
    const d = new Date(v);
    if (isNaN(d.getTime())) return v;
    return d.toISOString();
  }

  // ---------- versions ----------
  async function getFileVersions(role, siteUrl, serverRelPath) {
    const enc = spPath(serverRelPath);
    const path = `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')/Versions?$select=ID,VersionLabel,Url,Created,IsCurrentVersion,Size`;
    const data = await callRest(role, siteUrl, path);
    return (data.value || []).map(v => ({
      id: v.ID,
      label: v.VersionLabel,
      url: v.Url,
      created: v.Created,
      isCurrent: v.IsCurrentVersion,
      size: Number(v.Size || 0),
    }));
  }

  // Download a historical version's bytes via the _api (CORS-safe). The version's
  // physical _vti_history path is NOT resolvable through GetFileByServerRelativePath
  // ("file does not exist"); instead index the CURRENT file's Versions collection by
  // version ID and read its media stream: .../Versions(<id>)/$value.
  async function downloadFileVersion(role, siteUrl, serverRelPath, versionId) {
    const enc = spPath(serverRelPath);
    const res = await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')/Versions(${versionId})/$value`, { raw: true });
    return res.blob();
  }

  // ---------- delete ----------
  // Recycle instead of a hard delete, so every file SPOMove removes (sync deletes, move
  // sources, replaced targets) can be restored from the site recycle bin.
  async function deleteFile(role, siteUrl, serverRelPath) {
    const digest = await getFormDigest(role, siteUrl);
    const enc = spPath(serverRelPath);
    await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')/recycle()`, { method: 'POST', digest });
  }

  // Return { modified, size } if the target file exists, else null. Used for
  // overwrite/only-newer decisions before transferring.
  async function getFileInfoIfExists(role, siteUrl, serverRelPath) {
    const enc = spPath(serverRelPath);
    try {
      const data = await callRest(role, siteUrl, `/_api/web/GetFileByServerRelativePath(decodedurl='${enc}')?$select=TimeLastModified,Length,Exists`);
      if (data && data.Exists === false) return null;
      return { modified: data.TimeLastModified, size: Number(data.Length || 0) };
    } catch (e) {
      if (e.status === 404) return null;
      throw e;
    }
  }

  // ---------- server-side copy / move (same tenant only) ----------
  // Uses SP.MoveCopyUtil so bytes never travel through the browser.
  // Only KeepBoth and ShouldBypassSharedLocks are sent: most tenants reject
  // ResetAuthorAndEditorOnCopy / RetainEditorAndModifiedOnMove. Author/Editor are set
  // afterwards. A move does not bypass shared locks, so a file open in co-authoring
  // is not pulled away from its editors (it fails and can be retried).
  async function serverSideTransfer(role, contextSiteUrl, srcAbsoluteUrl, destAbsoluteUrl, { move = false, overwrite = true, keepBoth = false } = {}) {
    const digest = await getFormDigest(role, contextSiteUrl);
    const method = move ? 'MoveFileByPath' : 'CopyFileByPath';
    const options = { KeepBoth: keepBoth, ShouldBypassSharedLocks: !move };
    const body = JSON.stringify({
      srcPath: { DecodedUrl: srcAbsoluteUrl },
      destPath: { DecodedUrl: destAbsoluteUrl },
      overwrite: !!overwrite,
      options,
    });
    await callRest(role, contextSiteUrl, `/_api/SP.MoveCopyUtil.${method}()`, {
      method: 'POST', digest,
      headers: { 'Content-Type': 'application/json;odata=nometadata' },
      body,
    });
  }

  // Reset Author/Editor to a given user and Created/Modified to now. Used after a
  // server-side copy when the user chose NOT to preserve metadata (the copy would
  // otherwise carry over the original author/editor).
  async function resetFileMetadata(role, siteUrl, serverRelPath, upn) {
    const now = new Date().toISOString();
    const claim = upn ? buildClaimLoginName(upn) : null;
    const meta = {
      created: now,
      modified: now,
      author: claim ? { loginName: claim } : null,
      editor: claim ? { loginName: claim } : null,
    };
    await setFileMetadata(role, siteUrl, serverRelPath, meta, { includeAuthorEditor: !!claim });
  }

  // ---------- take ownership (SharePoint admin) ----------
  function xmlEscape(s) {
    return String(s || '').replace(/[<>&"']/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[c]));
  }
  function buildClaimLoginName(upn) {
    if (!upn) throw new Error('UPN required to build claim login name');
    if (/^i:0/i.test(upn)) return upn;
    return `i:0#.f|membership|${upn}`;
  }

  // ---------- tenant admin API (SharePoint Administrator role) ----------
  function adminOpts(role) {
    return { label: `SPO-admin[${role}]`, token: () => Auth.getSpoAdminToken(role) };
  }
  function adminHost(role) {
    const conn = Auth.getConnection(role);
    if (!conn || !conn.adminHost) throw new Error('SharePoint admin host not detected for this tenant');
    return conn.adminHost;
  }

  async function getAdminFormDigest(role) {
    const data = await Http.json(`https://${adminHost(role)}/_api/contextinfo`, {
      ...adminOpts(role), method: 'POST', headers: { Accept: 'application/json;odata=verbose' },
    });
    const value = data && ((data.d && data.d.GetContextWebInformation && data.d.GetContextWebInformation.FormDigestValue) || data.FormDigestValue);
    if (!value) throw new Error('Could not obtain admin form digest');
    return value;
  }

  // Every site collection in the tenant, regardless of the user's own site permissions.
  // Pages through GetSitePropertiesFromSharePointByFilters with the StartIndex cursor.
  async function enumerateAllSites(role, { includePersonalSites = false, onProgress } = {}) {
    const endpoint = `https://${adminHost(role)}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/GetSitePropertiesFromSharePointByFilters`;
    const digest = await getAdminFormDigest(role);
    const sites = [];
    let startIndex = null;
    let guard = 0;
    while (guard++ < 1000) {
      const speFilter = {
        __metadata: { type: 'Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerableFilter' },
        IncludePersonalSite: includePersonalSites ? 1 : 0,
        StartIndex: startIndex,
        IncludeDetail: false,
        GroupIdDefined: 0,
        Template: '',
        Filter: '',
      };
      const data = await Http.json(endpoint, {
        ...adminOpts(role),
        method: 'POST',
        headers: { 'X-RequestDigest': digest, 'Content-Type': 'application/json;odata=verbose', Accept: 'application/json;odata=verbose' },
        body: JSON.stringify({ speFilter }),
      });
      const d = (data && data.d && (data.d.GetSitePropertiesFromSharePointByFilters || data.d)) || data || {};
      const batch = d.results || d.value || [];
      for (const s of batch) {
        if (!s.Url) continue;
        sites.push({ url: trimUrl(s.Url), title: s.Title || s.Url, template: s.Template || '', storageUsage: Number(s.StorageUsageCurrent || s.StorageUsage || 0) });
      }
      if (onProgress) onProgress(sites.length);
      const nextIdx = d.NextStartIndexFromSharePoint;
      if (nextIdx == null || nextIdx === '' || nextIdx === '0') break;
      startIndex = String(nextIdx);
    }
    const seen = new Set();
    const unique = sites.filter(s => { const k = s.url.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
    unique.sort((a, b) => String(a.title || '').localeCompare(String(b.title || '')));
    return unique;
  }

  async function setSiteAdmin(role, siteUrl, claimLoginName, isSiteAdmin) {
    const conn = Auth.getConnection(role);
    if (!conn || !MoveRules.isTenantSpoHost(hostOf(siteUrl), conn.tenantHost)) throw new Error(`${siteUrl} is not a site of the ${role} tenant`);
    const url = `https://${adminHost(role)}/_vti_bin/client.svc/ProcessQuery`;
    const digest = await getAdminFormDigest(role).catch(() => null);
    const xml = `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SPOMove">` +
      `<Actions><Method Name="SetSiteAdmin" Id="1" ObjectPathId="2"><Parameters>` +
      `<Parameter Type="String">${xmlEscape(siteUrl)}</Parameter>` +
      `<Parameter Type="String">${xmlEscape(claimLoginName)}</Parameter>` +
      `<Parameter Type="Boolean">${isSiteAdmin ? 'true' : 'false'}</Parameter>` +
      `</Parameters></Method></Actions>` +
      `<ObjectPaths><Constructor Id="2" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
    const headers = { 'Content-Type': 'text/xml', Accept: '*/*' };
    if (digest) headers['X-RequestDigest'] = digest;
    Log.info(`SetSiteAdmin(${siteUrl}, ${claimLoginName}, ${isSiteAdmin})`);
    const res = await Http.request(url, { ...adminOpts(role), method: 'POST', headers, body: xml });
    const bodyJson = await res.json().catch(() => null);
    if (Array.isArray(bodyJson) && bodyJson[0] && bodyJson[0].ErrorInfo) {
      const ei = bodyJson[0].ErrorInfo;
      const e = new Error(`SetSiteAdmin failed: ${ei.ErrorMessage || 'unknown'}`);
      // -2147024891 = E_ACCESSDENIED: the tenant admin API refused it.
      if (String(ei.ErrorCode) === '-2147024891' || /unauthorized/i.test(ei.ErrorMessage || '')) e.status = 403;
      throw e;
    }
    return true;
  }

  window.SharePoint = {
    serverRelativeOf, hostOf, folderOf, nameOf, joinPath, trimUrl, spPath,
    validateConnection, getWeb, getDocumentLibraries, getFolderChildren, enumerateFiles, enumerateAllSites,
    createFolder, downloadFile, uploadFile,
    getFileMetadata, setFileMetadata, resetFileMetadata, ensureUser, getFileVersions, downloadFileVersion,
    getFileInfoIfExists, deleteFile, serverSideTransfer,
    getThrottleState,
    setSiteAdmin, buildClaimLoginName,
  };
})();
