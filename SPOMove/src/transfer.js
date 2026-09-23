// Transfer engine. Orchestrates a job end-to-end: enumerate source files, then
// copy/move each to the target with the chosen options, checkpointing progress
// so an interrupted run can resume. Same-tenant transfers use SharePoint's
// server-side MoveCopyUtil (no bytes through the browser); cross-tenant and
// version-preserving transfers stream bytes download->upload.
(() => {
  const SP = () => window.SharePoint;
  const running = new Map(); // jobId -> { cancelToken, promise }

  function rolesFor(job) {
    const sameTenant = job.config.mode === 'same';
    return { sourceRole: 'source', targetRole: sameTenant ? 'source' : 'target', sameTenant };
  }

  function snapshot(job, current) {
    const s = job.stats;
    return {
      status: job.status,
      itemsTotal: s.total, itemsDone: s.done, failed: s.failed, skipped: s.skipped,
      bytesTotal: s.bytesTotal, bytesDone: s.bytesDone,
      currentFile: current || '',
    };
  }

  async function run(jobId, callbacks = {}) {
    if (running.has(jobId)) return running.get(jobId).promise;
    let job = Jobs.get(jobId);
    if (!job) throw new Error('Job not found');
    const cancelToken = new Concurrency.CancelToken();
    const speed = { lastTs: Date.now(), lastBytes: job.stats.bytesDone || 0, bps: 0 };
    const ctx = { cancelToken, callbacks, speed, lastSave: 0, lastItemSave: 0, items: null };
    const promise = _run(job, ctx).catch(async e => {
      Log.err(`Job ${jobId} failed:`, e);
      if (ctx.items) { try { await Jobs.saveItems(jobId, ctx.items); } catch {} }
      job = Jobs.get(jobId) || job;
      if (job.status !== 'cancelled') { job.status = 'failed'; job.lastError = e.message; }
      job.finishedAt = Date.now();
      Jobs.save(job);
      _emit(ctx, job);
      throw e;
    }).finally(() => running.delete(jobId));
    running.set(jobId, { cancelToken, promise });
    return promise;
  }

  function pause(jobId) { running.get(jobId)?.cancelToken.pause(); }
  function resumeRun(jobId) { running.get(jobId)?.cancelToken.resume(); }
  function cancel(jobId) { running.get(jobId)?.cancelToken.cancel(); }
  function isRunning(jobId) { return running.has(jobId); }

  function _emit(ctx, job, current) {
    // Rolling speed estimate.
    const now = Date.now();
    const dt = (now - ctx.speed.lastTs) / 1000;
    if (dt >= 0.5) {
      const db = job.stats.bytesDone - ctx.speed.lastBytes;
      ctx.speed.bps = db / dt;
      ctx.speed.lastTs = now;
      ctx.speed.lastBytes = job.stats.bytesDone;
    }
    const snap = snapshot(job, current);
    snap.speedBps = ctx.speed.bps;
    const remainingBytes = Math.max(0, snap.bytesTotal - snap.bytesDone);
    snap.etaSec = ctx.speed.bps > 0 ? Math.round(remainingBytes / ctx.speed.bps) : null;
    ctx.callbacks.onProgress?.(snap, job);
  }

  function _saveThrottled(ctx, job, force = false) {
    const now = Date.now();
    if (force || now - ctx.lastSave > 1200) {
      Jobs.save(job);
      ctx.lastSave = now;
    }
  }

  // Item-status checkpoint to IndexedDB. Throttled hard because it rewrites the whole
  // (potentially huge) item array; on resume at most this interval of work is redone.
  async function _saveItemsThrottled(ctx, job, force = false) {
    const now = Date.now();
    if (force || now - ctx.lastItemSave > 10000) {
      ctx.lastItemSave = now;
      try { await Jobs.saveItems(job.id, job.items); }
      catch (e) { Log.warn(`Item checkpoint failed: ${e.message}`); }
    }
  }

  async function _run(job, ctx) {
    const roles = rolesFor(job);
    const cfg = job.config;

    // --- validate both connections ---
    if (!Auth.isConnected(roles.sourceRole)) throw new Error('Source tenant is not connected');
    if (roles.targetRole === 'target' && !Auth.isConnected('target')) throw new Error('Target tenant is not connected');

    const srcCheck = await SP().validateConnection(roles.sourceRole, cfg.source.siteUrl);
    if (!srcCheck.ok) throw new Error(`Source site not reachable: ${srcCheck.error}`);
    const dstCheck = await SP().validateConnection(roles.targetRole, cfg.target.siteUrl);
    if (!dstCheck.ok) throw new Error(`Target site not reachable: ${dstCheck.error}`);

    // --- re-run (re-scan + diff), resume (load), or first-run (enumerate) ---
    if (job.rerun) {
      const mode = job.rerun;
      job.rerun = null;
      const prevItems = await Jobs.loadItems(job.id).catch(() => []);
      job.status = 'enumerating';
      Jobs.save(job);
      _emit(ctx, job);
      Log.info(`Re-run (${mode}): re-scanning ${cfg.source.folder} ...`);
      const files = await SP().enumerateFiles(roles.sourceRole, cfg.source.siteUrl, cfg.source.folder, {
        cancelToken: ctx.cancelToken,
        onProgress: (found) => ctx.callbacks.onEnumerate?.(found),
      });
      job.items = buildRerunItems(mode, prevItems, files, cfg);
      ctx.items = job.items;
      if (cfg.includeVersions) await enrichVersions(job, roles, ctx);
      recomputeStats(job);
      await Jobs.saveItems(job.id, job.items);
      Jobs.save(job);
      const np = job.items.filter(i => i.status === 'pending' && !i.op).length;
      const nd = job.items.filter(i => i.op === 'delete').length;
      Log.info(`Re-run (${mode}): ${job.items.length} source item(s) — ${np} to (re)transfer, ${nd} to delete from target, ${job.stats.skipped} unchanged.`);
    } else {
      if (!job.items) job.items = await Jobs.loadItems(job.id);
      if (!job.items.length) {
        job.status = 'enumerating';
        Jobs.save(job);
        _emit(ctx, job);
        Log.info(`Enumerating files under ${cfg.source.folder} ...`);
        const files = await SP().enumerateFiles(roles.sourceRole, cfg.source.siteUrl, cfg.source.folder, {
          cancelToken: ctx.cancelToken,
          onProgress: (found) => ctx.callbacks.onEnumerate?.(found),
        });
        job.items = files.map(f => ({
          subPath: f.subPath,
          size: f.size,
          modified: f.modified,
          status: 'pending',
          error: null,
        }));
        ctx.items = job.items;
        // When preserving version history, read each file's versions up front so the
        // reported totals (size + count) and the progress bar reflect the ACTUAL data
        // moved. The compact list is cached on the item and reused at transfer time.
        if (cfg.includeVersions) await enrichVersions(job, roles, ctx);
        job.stats.total = job.items.length;
        job.stats.bytesTotal = job.items.reduce((a, b) => a + itemBytes(b), 0);
        job.stats.versionsTotal = job.items.reduce((a, b) => a + (b.versionCount || 0), 0);
        await Jobs.saveItems(job.id, job.items);
        Jobs.save(job);
        const vNote = job.stats.versionsTotal ? ` + ${job.stats.versionsTotal} version(s)` : '';
        Log.info(`Found ${job.items.length} files${vNote} (${fmtBytes(job.stats.bytesTotal)}).`);
      } else {
        ctx.items = job.items;
        recomputeStats(job);
      }
    }

    if (job.stats.total === 0) {
      job.status = 'completed';
      job.finishedAt = Date.now();
      Jobs.save(job);
      _emit(ctx, job);
      return job;
    }

    // --- run ---
    job.status = 'running';
    job.startedAt = job.startedAt || Date.now();
    job.lastError = null;
    Jobs.save(job);
    _emit(ctx, job);

    const hostSrc = SP().hostOf(cfg.source.siteUrl);
    const ensureFolder = makeFolderEnsurer(cfg, roles, ctx);
    const pending = job.items.filter(it => it.status !== 'done' && it.status !== 'skipped' && it.status !== 'deleted');

    await Concurrency.pmap(pending, async (item) => {
      ctx.cancelToken.throwIfCancelled();
      await ctx.cancelToken.waitIfPaused();
      item.attempts = (item.attempts || 0) + 1;
      try {
        // Race against cancellation so a click stops the job promptly even while an item
        // is stuck in a long throttle wait or transfer (the orphaned op self-terminates
        // at the next cancellation checkpoint inside transferItem).
        const res = await Concurrency.raceCancel(
          transferItem(job, roles, item, { ensureFolder, hostSrc, ctx }),
          ctx.cancelToken.signal
        );
        if (res && res.deleted) {
          item.status = 'deleted';
          item.error = null;
          item.skipReason = null;
          job.stats.deleted = (job.stats.deleted || 0) + 1;
        } else if (res && res.skipped) {
          item.status = 'skipped';
          item.error = null;
          item.skipReason = res.reason || 'Skipped';
          job.stats.skipped += 1;
        } else {
          item.status = 'done';
          item.error = null;
          item.skipReason = null;
          job.stats.done += 1;
        }
        job.stats.bytesDone += itemBytes(item);
      } catch (e) {
        if (/cancelled/i.test(e.message)) throw e;
        item.status = 'failed';
        item.error = e.message;
        item.skipReason = null;
        job.stats.failed += 1;
        Log.warn(`Failed ${item.subPath} (attempt ${item.attempts}): ${e.message}`);
      }
      _saveThrottled(ctx, job);
      await _saveItemsThrottled(ctx, job);
      _emit(ctx, job, item.subPath);
    }, { concurrency: cfg.concurrency || 4, cancelToken: ctx.cancelToken });

    // --- finalize ---
    if (ctx.cancelToken.cancelled) {
      job.status = 'cancelled';
    } else {
      job.status = job.stats.failed > 0 ? 'completed_with_errors' : 'completed';
    }
    job.finishedAt = Date.now();
    await _saveItemsThrottled(ctx, job, true);
    Jobs.save(job);
    _emit(ctx, job);
    Log.info(`Job ${job.id} ${job.status}: ${job.stats.done} done, ${job.stats.failed} failed.`);
    return job;
  }

  // Bytes a single item accounts for: its current content plus any version history
  // that will be replayed (0 unless includeVersions enriched the item).
  function itemBytes(item) { return (item.size || 0) + (item.versionBytes || 0); }

  // Read each file's non-current version history in parallel and cache a compact list
  // on the item, so totals are accurate before the transfer starts and the transfer
  // reuses the list instead of fetching it again.
  async function enrichVersions(job, roles, ctx) {
    const cfg = job.config;
    const SPm = SP();
    const base = SPm.trimUrl(cfg.source.folder);
    const targets = job.items.filter(i => i.status === 'pending' && !i.op);
    if (!targets.length) return;
    let done = 0;
    Log.info(`Reading version history for ${targets.length} file(s) ...`);
    await Concurrency.pmap(targets, async (item) => {
      ctx.cancelToken.throwIfCancelled();
      try {
        const versions = (await SPm.getFileVersions(roles.sourceRole, cfg.source.siteUrl, `${base}/${item.subPath}`))
          .filter(v => !v.isCurrent)
          .sort((a, b) => new Date(a.created) - new Date(b.created))
          .map(v => ({ id: v.id, label: v.label, created: v.created, size: v.size }));
        item.versions = versions;
        item.versionCount = versions.length;
        item.versionBytes = versions.reduce((a, v) => a + (v.size || 0), 0);
      } catch (e) {
        // Leave item.versions undefined so transferItem fetches it at transfer time.
        item.versionCount = 0;
        item.versionBytes = 0;
        Log.warn(`Could not read versions for ${item.subPath}: ${e.message}`);
      }
      if (++done % 25 === 0 || done === targets.length) ctx.callbacks.onVersions?.(done, targets.length);
    }, { concurrency: cfg.concurrency || 4, cancelToken: ctx.cancelToken });
  }

  function recomputeStats(job) {
    const s = { total: job.items.length, done: 0, failed: 0, skipped: 0, deleted: 0, bytesTotal: 0, bytesDone: 0, versionsTotal: 0 };
    for (const it of job.items) {
      const tb = itemBytes(it);
      s.bytesTotal += tb;
      s.versionsTotal += it.versionCount || 0;
      if (it.status === 'done') { s.done++; s.bytesDone += tb; }
      else if (it.status === 'skipped') { s.skipped++; s.bytesDone += tb; }
      else if (it.status === 'deleted') { s.deleted++; }
      else if (it.status === 'failed') { it.status = 'pending'; } // retry failures on resume
    }
    job.stats = s;
  }

  // Build the item list for a re-run by diffing a fresh source scan against the previous run.
  // 'full' re-transfers everything currently in the source. 'incremental' transfers only new +
  // changed files, carries unchanged files as skipped, and (for copy jobs) queues a target
  // delete for files that were transferred before but are now gone from the source. Move jobs
  // never queue deletes (their transferred files are meant to be absent from the source).
  function buildRerunItems(mode, prevItems, files, cfg) {
    const prevByPath = new Map((prevItems || []).map(i => [i.subPath, i]));
    const seen = new Set();
    const out = [];
    for (const f of files) {
      seen.add(f.subPath);
      const base = { subPath: f.subPath, size: f.size, modified: f.modified, status: 'pending', error: null };
      if (mode === 'full') { out.push(base); continue; }
      const prev = prevByPath.get(f.subPath);
      if (!prev) { out.push(base); continue; }                               // new file
      if (isSourceNewer(f.modified, prev.modified)) { out.push(base); continue; } // changed
      out.push({ ...base, status: 'skipped', skipReason: 'Unchanged since last run' });
    }
    if (mode === 'incremental' && cfg.operation !== 'move') {
      for (const p of (prevItems || [])) {
        if (seen.has(p.subPath)) continue;
        if (p.status === 'done') out.push({ subPath: p.subPath, size: 0, modified: p.modified, status: 'pending', error: null, op: 'delete' });
      }
    }
    return out;
  }

  // Build a folder-ensurer that creates each destination folder exactly once and always
  // creates parents before children. Files in a new folder await the SAME creation
  // promise, so nothing is transferred before its folder exists (folders are on the
  // critical path). The target base folder is assumed to already exist.
  function makeFolderEnsurer(cfg, roles, ctx) {
    const SPm = SP();
    const base = SPm.trimUrl(cfg.target.folder);
    const cache = new Map();
    function ensure(folderPath) {
      const path = SPm.trimUrl(folderPath);
      if (path.length <= base.length) return Promise.resolve();
      if (cache.has(path)) return cache.get(path);
      const p = (async () => {
        await ensure(SPm.folderOf(path));
        ctx.cancelToken.throwIfCancelled();
        await SPm.createFolder(roles.targetRole, cfg.target.siteUrl, path);
      })();
      cache.set(path, p);
      p.catch(() => cache.delete(path)); // let a later file retry a failed create
      return p;
    }
    return ensure;
  }

  async function transferItem(job, roles, item, { ensureFolder, hostSrc, ctx }) {
    const cfg = job.config;
    const SPm = SP();
    const srcSite = cfg.source.siteUrl;
    const dstSite = cfg.target.siteUrl;
    const name = SPm.nameOf(item.subPath);
    const srcServerRel = `${SPm.trimUrl(cfg.source.folder)}/${item.subPath}`;
    const destServerRel = `${SPm.trimUrl(cfg.target.folder)}/${item.subPath}`;
    const destFolder = SPm.folderOf(destServerRel);

    ctx.cancelToken.throwIfCancelled();

    // Incremental re-run: the source object was removed since the last run — remove it from
    // the target too (a 404 means it is already gone, which is fine).
    if (item.op === 'delete') {
      try { await SPm.deleteFile(roles.targetRole, dstSite, destServerRel); }
      catch (e) { if (e.status !== 404) throw e; }
      return { deleted: true };
    }

    await ensureFolder(destFolder);

    const byteMode = !roles.sameTenant || cfg.includeVersions;

    // Existence-aware handling. onlyNewer: skip when the target exists and the source
    // is not newer (timestamps are UTC from SharePoint, so the compare is timezone-safe).
    // overwrite+byteMode: delete the existing target first so replayed version history
    // starts clean instead of stacking new versions onto the old target file.
    if (cfg.overwrite && (cfg.onlyNewer || byteMode)) {
      let existing = null;
      try { existing = await SPm.getFileInfoIfExists(roles.targetRole, dstSite, destServerRel); }
      catch (e) { Log.warn(`Could not check target for ${name}: ${e.message}`); }
      if (existing) {
        if (cfg.onlyNewer && !isSourceNewer(item.modified, existing.modified)) {
          return { skipped: true, reason: 'Target already exists and is newer or the same age (only-newer is on)' };
        }
        if (byteMode) {
          try { await SPm.deleteFile(roles.targetRole, dstSite, destServerRel); }
          catch (e) { Log.warn(`Could not delete existing target ${name}: ${e.message}`); }
        }
      }
    }

    if (!byteMode) {
      // Fast path: server-side copy/move within the same tenant.
      const hostDst = SPm.hostOf(dstSite);
      const srcAbs = `https://${hostSrc}${srcServerRel}`;
      const destAbs = `https://${hostDst}${destServerRel}`;
      const move = cfg.operation === 'move';
      const doTransfer = () => SPm.serverSideTransfer(roles.sourceRole, srcSite, srcAbs, destAbs, { move, overwrite: cfg.overwrite });
      ctx.cancelToken.throwIfCancelled();
      try {
        await doTransfer();
      } catch (e) {
        // MoveCopyUtil returns 403/423 when it refuses to clobber an existing destination
        // (typically on a full re-run overwriting files it created before). If overwrite is
        // on and the target really exists, delete it and transfer fresh.
        if (!(cfg.overwrite && (e.status === 403 || e.status === 423))) throw e;
        let existed = null;
        try { existed = await SPm.getFileInfoIfExists(roles.targetRole, dstSite, destServerRel); } catch {}
        if (!existed) throw accessError(e, name);
        await SPm.deleteFile(roles.targetRole, dstSite, destServerRel).catch(de => { throw accessError(de, name); });
        ctx.cancelToken.throwIfCancelled();
        await doTransfer().catch(re => { throw accessError(re, name); });
      }
      // A copy carries over the original Author/Editor and can reset Created/Modified.
      // Re-sync the original metadata when preserving it, or reset to the current
      // user/now when not (the old ResetAuthorAndEditorOnCopy flag is unsupported).
      if (cfg.operation === 'copy') {
        try {
          if (cfg.includeMetadata) {
            const meta = await SPm.getFileMetadata(roles.sourceRole, srcSite, srcServerRel);
            await SPm.setFileMetadata(roles.targetRole, dstSite, destServerRel, meta, { includeAuthorEditor: true });
          } else {
            const upn = Auth.getConnection(roles.targetRole)?.upn;
            await SPm.resetFileMetadata(roles.targetRole, dstSite, destServerRel, upn);
          }
        } catch (e) { Log.warn(`Metadata update failed for ${name}: ${e.message}`); }
      }
      return;
    }

    // Byte path: optional version replay, then current content, then metadata.
    if (cfg.includeVersions) {
      let versions = item.versions;
      if (!versions) {
        try {
          versions = (await SPm.getFileVersions(roles.sourceRole, srcSite, srcServerRel))
            .filter(v => !v.isCurrent).sort((a, b) => new Date(a.created) - new Date(b.created));
        } catch (e) { Log.warn(`Could not read versions for ${name}: ${e.message}`); versions = []; }
      }
      for (const v of versions) {
        try {
          const blob = await SPm.downloadFileVersion(roles.sourceRole, srcSite, srcServerRel, v.id);
          await SPm.uploadFile(roles.targetRole, dstSite, destFolder, name, blob, { overwrite: true, cancelToken: ctx.cancelToken });
        } catch (e) { Log.warn(`Version ${v.label} of ${name} skipped: ${e.message}`); }
      }
    }

    const blob = await SPm.downloadFile(roles.sourceRole, srcSite, srcServerRel);
    try {
      await SPm.uploadFile(roles.targetRole, dstSite, destFolder, name, blob, {
        overwrite: cfg.overwrite, cancelToken: ctx.cancelToken,
        onProgress: (sent) => { /* per-file byte progress folds into item.size on completion */ },
      });
    } catch (e) { throw accessError(e, name); }

    if (cfg.includeMetadata) {
      try {
        const meta = await SPm.getFileMetadata(roles.sourceRole, srcSite, srcServerRel);
        await SPm.setFileMetadata(roles.targetRole, dstSite, destServerRel, meta, { includeAuthorEditor: true });
      } catch (e) { Log.warn(`Metadata sync failed for ${name}: ${e.message}`); }
    }

    if (cfg.operation === 'move') {
      ctx.cancelToken.throwIfCancelled();
      await SPm.deleteFile(roles.sourceRole, srcSite, srcServerRel);
    }
  }

  // Whether the source file is strictly newer than the target. Unknown/unparseable
  // timestamps default to true (allow the overwrite) rather than silently skipping.
  function isSourceNewer(srcModified, tgtModified) {
    if (!srcModified || !tgtModified) return true;
    const s = new Date(srcModified).getTime();
    const t = new Date(tgtModified).getTime();
    if (isNaN(s) || isNaN(t)) return true;
    return s > t;
  }

  // Wrap a raw 403/423 with an actionable hint: the target file is protected, not a tool bug.
  function accessError(e, name) {
    if (e && (e.status === 403 || e.status === 423)) {
      const ne = new Error(`Access denied for ${name} — the target file is likely checked out, open/locked by a user, or a declared record / on hold in the destination. (${e.message})`);
      ne.status = e.status;
      return ne;
    }
    return e;
  }

  function fmtBytes(n) {
    if (!n) return '0 B';
    const u = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(n) / Math.log(1024));
    return `${(n / Math.pow(1024, i)).toFixed(i ? 1 : 0)} ${u[i]}`;
  }

  window.Transfer = { run, pause, resumeRun, cancel, isRunning, fmtBytes };
})();
