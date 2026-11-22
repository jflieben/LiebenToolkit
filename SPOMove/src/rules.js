// Pure decision logic for SPOMove, kept apart so it can be unit tested:
// source/target overlap, sync (re-run) planning, existing-target handling,
// SharePoint host checks and a guard that keeps tokens out of stored job data.
(() => {
  function normPath(p) {
    let s = String(p == null ? '' : p);
    try { s = decodeURIComponent(s); } catch { /* keep as is */ }
    return s.replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
  }

  function hostOf(url) {
    try { return new URL(url).host.toLowerCase(); } catch { return ''; }
  }

  // SharePoint paths are case-insensitive.
  function isSameOrUnder(path, folder) {
    const p = normPath(path);
    const f = normPath(folder);
    if (!f) return false;
    return p === f || p.startsWith(`${f}/`);
  }

  // loc: { tenantId, siteUrl, folder } with folder server-relative. Returns a reason or null.
  function overlapError(src, tgt) {
    if (!src || !tgt || !src.folder || !tgt.folder) return null;
    if (src.tenantId && tgt.tenantId && src.tenantId !== tgt.tenantId) return null;
    if (hostOf(src.siteUrl) !== hostOf(tgt.siteUrl)) return null;
    const s = normPath(src.folder);
    const t = normPath(tgt.folder);
    if (s === t) return 'Source and target are the same folder.';
    if (t.startsWith(`${s}/`)) return 'The target folder is inside the source folder. Pick a target outside the source.';
    if (s.startsWith(`${t}/`)) return 'The source folder is inside the target folder. Pick a target outside the source.';
    return null;
  }

  function sameFile(srcAbsUrl, dstAbsUrl) {
    return hostOf(srcAbsUrl) === hostOf(dstAbsUrl) && normPath(pathOf(srcAbsUrl)) === normPath(pathOf(dstAbsUrl));
  }
  function pathOf(url) { try { return new URL(url).pathname; } catch { return String(url || ''); } }

  // Unknown or unparseable timestamps count as newer, so nothing is skipped silently.
  function isSourceNewer(srcModified, tgtModified) {
    if (!srcModified || !tgtModified) return true;
    const s = new Date(srcModified).getTime();
    const t = new Date(tgtModified).getTime();
    if (isNaN(s) || isNaN(t)) return true;
    return s > t;
  }

  // What to do when the target may already hold the file.
  // 'proceed' | 'skip-exists' | 'skip-not-newer' | 'replace' (delete the target first, byte mode)
  function existingTargetAction({ overwrite, onlyNewer } = {}, existing, srcModified, byteMode) {
    if (!existing) return 'proceed';
    if (!overwrite) return 'skip-exists';
    if (onlyNewer && !isSourceNewer(srcModified, existing.modified)) return 'skip-not-newer';
    return byteMode ? 'replace' : 'proceed';
  }

  // Overwrite flag for one upload: the first upload of a file follows the Overwrite option,
  // later uploads (more versions, then the current content) replace the file this run created.
  function uploadOverwrite(overwriteOption, createdThisRun) {
    return !!createdThisRun || !!overwriteOption;
  }

  // Re-run plan from the previous item list and a fresh source scan.
  // 'full' re-transfers everything in the source. 'incremental' transfers new and changed
  // files, carries unchanged ones as skipped and, for copy jobs, proposes target deletes for
  // files this job transferred before that are gone from the source now. Deletes are only
  // proposed when every source folder could be listed: a folder that failed to list would
  // otherwise look empty and its files would be deleted from the target.
  function planRerun(mode, prevItems, files, { operation, unlistedFolders = [] } = {}) {
    const prevByPath = new Map((prevItems || []).map((i) => [i.subPath, i]));
    const seen = new Set();
    const items = [];
    for (const f of files || []) {
      seen.add(f.subPath);
      const base = { subPath: f.subPath, size: f.size, modified: f.modified, status: 'pending', error: null };
      if (mode === 'full') { items.push(base); continue; }
      const prev = prevByPath.get(f.subPath);
      if (!prev || isSourceNewer(f.modified, prev.modified)) { items.push(base); continue; }
      items.push({ ...base, status: 'skipped', skipReason: 'Unchanged since last run', transferred: wasTransferred(prev) });
    }
    const candidates = [];
    if (mode === 'incremental' && operation !== 'move') {
      for (const p of prevItems || []) {
        if (!seen.has(p.subPath) && wasTransferred(p)) candidates.push({ subPath: p.subPath, modified: p.modified });
      }
    }
    const blocked = (unlistedFolders || []).length > 0;
    return {
      items,
      deletes: blocked ? [] : candidates,
      blockedDeletes: blocked ? candidates.length : 0,
    };
  }

  // A file this job put on the target (or a sync delete that has not happened yet).
  function wasTransferred(item) {
    if (!item) return false;
    if (item.op === 'delete') return item.status !== 'deleted';
    return item.status === 'done' || item.transferred === true;
  }

  function deleteItems(deletes) {
    return (deletes || []).map((d) => ({ subPath: d.subPath, size: 0, modified: d.modified, status: 'pending', error: null, op: 'delete' }));
  }

  // Share of the files this job put on the target that a sync would delete.
  function deleteShare(deleteCount, prevItems) {
    const base = (prevItems || []).filter(wasTransferred).length;
    return base ? deleteCount / base : 0;
  }

  // A site URL from search, a paste or an API response only gets a token when it is a
  // SharePoint host of the connected tenant (incl. -my, -admin and multi-geo hosts).
  function isTenantSpoHost(host, tenantHost) {
    const h = String(host || '').toLowerCase();
    const m = /^([a-z0-9]+)\.sharepoint\.com$/.exec(String(tenantHost || '').toLowerCase());
    if (!m) return false;
    const re = new RegExp(`^${m[1]}([a-z]{3})?(-my|-admin)?\\.sharepoint\\.com$`);
    return re.test(h);
  }

  // Resume rule: a cached account may only stand in for a job's connection when both the
  // tenant id and the UPN are the job's.
  function accountMatches(account, loc) {
    if (!account || !loc || !loc.tenantId || !loc.upn) return false;
    return account.tenantId === loc.tenantId && String(account.username || '').toLowerCase() === String(loc.upn).toLowerCase();
  }

  const SECRET_KEY = /^(access_?token|refresh_?token|id_?token|secret|password|authorization)$/i;
  const JWT = /eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{4,}/;
  function containsSecret(value) {
    let found = false;
    const text = JSON.stringify(value, (k, v) => {
      if (k && SECRET_KEY.test(k)) found = true;
      return v;
    });
    return found || JWT.test(text || '') || /Bearer\s+[A-Za-z0-9\-._~+/]{20,}/.test(text || '');
  }

  window.MoveRules = {
    normPath, isSameOrUnder, overlapError, sameFile, isSourceNewer, existingTargetAction, uploadOverwrite,
    planRerun, deleteItems, deleteShare, wasTransferred, isTenantSpoHost, accountMatches, containsSecret,
  };
})();
