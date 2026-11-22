// Transfer engine. Runs a job end to end: enumerate the source, then copy or move each
// file to the target with the chosen options, checkpointing progress so an interrupted
// run can resume. Same tenant uses SharePoint's server-side MoveCopyUtil (no bytes through
// the browser); cross-tenant and version-preserving transfers stream download to upload.
(() => {
  const SP = () => window.SharePoint;
  const running = new Map(); // jobId -> { cancelToken, promise, roles }

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
    const promise = _run(job, ctx).catch(async (e) => {
      Log.err(`Job ${jobId} failed:`, e);
      if (ctx.items) { try { await Jobs.saveItems(jobId, ctx.items); } catch { /* best effort */ } }
      const latest = Jobs.get(jobId) || job;
      if (cancelToken.cancelled || /^Cancelled$/.test(e.message || '')) latest.status = 'cancelled';
      else { latest.status = 'failed'; latest.lastError = e.message; }
      latest.finishedAt = Date.now();
      Object.assign(job, latest);
      Jobs.save(job);
      _emit(ctx, job);
      throw e;
    }).finally(() => running.delete(jobId));
    running.set(jobId, { cancelToken, promise, roles: rolesFor(job) });
    return promise;
  }

  function pause(jobId) { running.get(jobId)?.cancelToken.pause(); }
  function resumeRun(jobId) { running.get(jobId)?.cancelToken.resume(); }
  function cancel(jobId) { running.get(jobId)?.cancelToken.cancel(); }
  function isRunning(jobId) { return running.has(jobId); }
  // Connection roles that running jobs depend on right now.
  function rolesInUse() {
    const out = new Set();
    for (const r of running.values()) { out.add(r.roles.sourceRole); out.add(r.roles.targetRole); }
    return out;
  }

  function _emit(ctx, job, current) {
    const now = Date.now();
    const dt = (now - ctx.speed.lastTs) / 1000;
    if (dt >= 0.5) {
      ctx.speed.bps = (job.stats.bytesDone - ctx.speed.lastBytes) / dt;
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

  // Rewrites the whole item array, so it is throttled hard; a resume redoes at most this interval.
  async function _saveItemsThrottled(ctx, job, force = false) {
    const now = Date.now();
    if (force || now - ctx.lastItemSave > 10000) {
      ctx.lastItemSave = now;
      try { await Jobs.saveItems(job.id, job.items); } catch (e) { Log.warn(`Item checkpoint failed: ${e.message}`); }
    }
  }

  async function _run(job, ctx) {
    const roles = rolesFor(job);
    const cfg = job.config;

    if (!Auth.isConnected(roles.sourceRole)) throw new Error('Source tenant is not connected');
    if (roles.targetRole === 'target' && !Auth.isConnected('target')) throw new Error('Target tenant is not connected');

    const overlap = MoveRules.overlapError(cfg.source, cfg.target);
    if (overlap) throw new Error(overlap);

    const srcCheck = await SP().validateConnection(roles.sourceRole, cfg.source.siteUrl);
    if (!srcCheck.ok) throw new Error(`Source site not reachable: ${srcCheck.error}`);
    const dstCheck = await SP().validateConnection(roles.targetRole, cfg.target.siteUrl);
    if (!dstCheck.ok) throw new Error(`Target site not reachable: ${dstCheck.error}`);

    if (job.rerun) {
      const mode = job.rerun;
      job.rerun = null;
      const prevItems = await Jobs.loadItems(job.id).catch(() => []);
      job.status = 'enumerating';
      Jobs.save(job);
      _emit(ctx, job);
      Log.info(`Re-run (${mode}): re-scanning ${cfg.source.folder} ...`);
      const scan = await scanSource(job, roles, ctx);
      const plan = MoveRules.planRerun(mode, prevItems, scan.files, { operation: cfg.operation, unlistedFolders: scan.failedFolders });
      job.items = plan.items;
      ctx.items = job.items;
      if (plan.blockedDeletes) {
        Log.warn(`Sync: ${scan.failedFolders.length} source folder(s) could not be listed, so nothing is deleted from the target in this run (${plan.blockedDeletes} file(s) would have been).`);
      }
      if (plan.deletes.length) {
        const share = MoveRules.deleteShare(plan.deletes.length, prevItems);
        const paths = plan.deletes.map((d) => `${SP().trimUrl(cfg.target.folder)}/${d.subPath}`);
        const ok = ctx.callbacks.confirmDeletes ? await ctx.callbacks.confirmDeletes(job, paths, share) : false;
        ctx.cancelToken.throwIfCancelled();
        if (ok) {
          job.items.push(...MoveRules.deleteItems(plan.deletes));
          Log.info(`Sync: deleting ${plan.deletes.length} target file(s) after confirmation.`);
        } else {
          Log.info(`Sync: kept ${plan.deletes.length} target file(s) that are gone from the source (not confirmed).`);
        }
      }
      if (cfg.includeVersions) await enrichVersions(job, roles, ctx);
      recomputeStats(job);
      await Jobs.saveItems(job.id, job.items);
      Jobs.save(job);
      const np = job.items.filter((i) => i.status === 'pending' && !i.op).length;
      const nd = job.items.filter((i) => i.op === 'delete').length;
      Log.info(`Re-run (${mode}): ${job.items.length} item(s), ${np} to transfer, ${nd} to delete from the target, ${job.stats.skipped} unchanged.`);
    } else {
      if (!job.items) job.items = await Jobs.loadItems(job.id);
      if (!job.items.length) {
        job.status = 'enumerating';
        Jobs.save(job);
        _emit(ctx, job);
        Log.info(`Enumerating files under ${cfg.source.folder} ...`);
        const scan = await scanSource(job, roles, ctx);
        job.items = scan.files.map((f) => ({ subPath: f.subPath, size: f.size, modified: f.modified, status: 'pending', error: null }));
        ctx.items = job.items;
        // With versions, read the history up front so totals and progress match the data
        // that really moves. The list is cached on the item for the transfer.
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

    const unlisted = (job.unlistedFolders || []).length > 0;
    if (job.stats.total === 0) {
      job.status = unlisted ? 'completed_with_errors' : 'completed';
      job.finishedAt = Date.now();
      Jobs.save(job);
      _emit(ctx, job);
      return job;
    }

    job.status = 'running';
    job.startedAt = job.startedAt || Date.now();
    job.lastError = null;
    Jobs.save(job);
    _emit(ctx, job);

    const hostSrc = SP().hostOf(cfg.source.siteUrl);
    const ensureFolder = makeFolderEnsurer(cfg, roles, ctx);
    const pending = job.items.filter((it) => it.status !== 'done' && it.status !== 'skipped' && it.status !== 'deleted');

    await Concurrency.pmap(pending, async (item) => {
      ctx.cancelToken.throwIfCancelled();
      await ctx.cancelToken.waitIfPaused();
      item.attempts = (item.attempts || 0) + 1;
      try {
        // Raced against cancellation so a click stops the job even during a long throttle
        // wait; the orphaned call stops at its next cancellation check.
        const res = await Concurrency.raceCancel(transferItem(job, roles, item, { ensureFolder, hostSrc, ctx }), ctx.cancelToken.signal);
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

    if (ctx.cancelToken.cancelled) job.status = 'cancelled';
    else job.status = job.stats.failed > 0 || unlisted ? 'completed_with_errors' : 'completed';
    job.finishedAt = Date.now();
    await _saveItemsThrottled(ctx, job, true);
    Jobs.save(job);
    _emit(ctx, job);
    Log.info(`Job ${job.id} ${job.status}: ${job.stats.done} done, ${job.stats.failed} failed${unlisted ? `, ${job.unlistedFolders.length} folder(s) not listed` : ''}.`);
    return job;
  }

  // The root must list. A subfolder that fails is recorded on the job (card and report)
  // and blocks sync deletes for this run.
  async function scanSource(job, roles, ctx) {
    const cfg = job.config;
    const scan = await SP().enumerateFiles(roles.sourceRole, cfg.source.siteUrl, cfg.source.folder, {
      cancelToken: ctx.cancelToken,
      onProgress: (found) => ctx.callbacks.onEnumerate?.(found),
    });
    const rootFail = scan.failedFolders.find((f) => MoveRules.normPath(f.folder) === MoveRules.normPath(cfg.source.folder));
    if (rootFail) throw new Error(`Could not list the source folder: ${rootFail.error}`);
    job.unlistedFolders = scan.failedFolders.map((f) => ({ folder: f.folder, error: f.error }));
    if (job.unlistedFolders.length) Log.warn(`${job.unlistedFolders.length} source folder(s) could not be listed; their files are not part of this run.`);
    return scan;
  }

  // Current content plus any version history that will be replayed.
  function itemBytes(item) { return (item.size || 0) + (item.versionBytes || 0); }

  async function enrichVersions(job, roles, ctx) {
    const cfg = job.config;
    const SPm = SP();
    const base = SPm.trimUrl(cfg.source.folder);
    const targets = job.items.filter((i) => i.status === 'pending' && !i.op);
    if (!targets.length) return;
    let done = 0;
    Log.info(`Reading version history for ${targets.length} file(s) ...`);
    await Concurrency.pmap(targets, async (item) => {
      ctx.cancelToken.throwIfCancelled();
      try {
        const versions = (await SPm.getFileVersions(roles.sourceRole, cfg.source.siteUrl, `${base}/${item.subPath}`))
          .filter((v) => !v.isCurrent)
          .sort((a, b) => new Date(a.created) - new Date(b.created))
          .map((v) => ({ id: v.id, label: v.label, created: v.created, size: v.size }));
        item.versions = versions;
        item.versionCount = versions.length;
        item.versionBytes = versions.reduce((a, v) => a + (v.size || 0), 0);
      } catch (e) {
        // transferItem fetches the list again at transfer time.
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
      else if (it.status === 'failed') { it.status = 'pending'; } // failures are retried on resume
    }
    job.stats = s;
  }

  // Creates each destination folder once, parents before children. Files in a new folder
  // wait for the same promise. The target base folder already exists.
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
      p.catch(() => cache.delete(path));
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
    const srcAbs = `https://${hostSrc}${srcServerRel}`;
    const destAbs = `https://${SPm.hostOf(dstSite)}${destServerRel}`;

    ctx.cancelToken.throwIfCancelled();
    if (MoveRules.sameFile(srcAbs, destAbs)) throw new Error('Source and target are the same file. Nothing was changed.');

    // Sync: the source file is gone, remove the copy this job made (404: already gone).
    if (item.op === 'delete') {
      try { await SPm.deleteFile(roles.targetRole, dstSite, destServerRel); }
      catch (e) { if (e.status !== 404) throw e; }
      return { deleted: true };
    }

    await ensureFolder(destFolder);

    const byteMode = !roles.sameTenant || cfg.includeVersions;

    // Byte mode always looks first: overwrite off skips an existing file before any version
    // is replayed, overwrite on removes it (to the recycle bin) so the history starts clean.
    if (byteMode || (cfg.overwrite && cfg.onlyNewer)) {
      let existing = null;
      try { existing = await SPm.getFileInfoIfExists(roles.targetRole, dstSite, destServerRel); }
      catch (e) { Log.warn(`Could not check the target for ${name}: ${e.message}`); }
      const action = MoveRules.existingTargetAction(cfg, existing, item.modified, byteMode);
      if (action === 'skip-exists') return { skipped: true, reason: 'Target file exists and overwrite is off' };
      if (action === 'skip-not-newer') return { skipped: true, reason: 'Target exists and is the same age or newer (only-newer is on)' };
      if (action === 'replace') {
        try { await SPm.deleteFile(roles.targetRole, dstSite, destServerRel); }
        catch (e) { Log.warn(`Could not remove the existing target ${name}: ${e.message}`); }
      }
    }

    if (!byteMode) {
      const move = cfg.operation === 'move';
      const doTransfer = () => SPm.serverSideTransfer(roles.sourceRole, srcSite, srcAbs, destAbs, { move, overwrite: cfg.overwrite });
      ctx.cancelToken.throwIfCancelled();
      try {
        await doTransfer();
      } catch (e) {
        // MoveCopyUtil answers 403/423 when it will not replace an existing destination. With
        // overwrite on and the target really there, remove it and transfer again.
        if (!(cfg.overwrite && (e.status === 403 || e.status === 423))) throw e;
        let existed = null;
        try { existed = await SPm.getFileInfoIfExists(roles.targetRole, dstSite, destServerRel); } catch { /* treat as unknown */ }
        if (!existed) throw accessError(e, name);
        await SPm.deleteFile(roles.targetRole, dstSite, destServerRel).catch((de) => { throw accessError(de, name); });
        ctx.cancelToken.throwIfCancelled();
        await doTransfer().catch((re) => { throw accessError(re, name); });
      }
      // A copy keeps Author/Editor but can reset Created/Modified: put the original back,
      // or reset to the current user and now when metadata is not preserved.
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
      return undefined;
    }

    // Byte path: version replay, then the current content, then metadata. The first upload
    // follows the Overwrite option; later ones replace the file this run just created.
    let created = false;
    if (cfg.includeVersions) {
      let versions = item.versions;
      if (!versions) {
        try {
          versions = (await SPm.getFileVersions(roles.sourceRole, srcSite, srcServerRel))
            .filter((v) => !v.isCurrent).sort((a, b) => new Date(a.created) - new Date(b.created));
        } catch (e) { Log.warn(`Could not read versions for ${name}: ${e.message}`); versions = []; }
      }
      for (const v of versions) {
        ctx.cancelToken.throwIfCancelled();
        try {
          const blob = await SPm.downloadFileVersion(roles.sourceRole, srcSite, srcServerRel, v.id);
          await SPm.uploadFile(roles.targetRole, dstSite, destFolder, name, blob, { overwrite: MoveRules.uploadOverwrite(cfg.overwrite, created), cancelToken: ctx.cancelToken });
          created = true;
        } catch (e) { Log.warn(`Version ${v.label} of ${name} skipped: ${e.message}`); }
      }
    }

    const blob = await SPm.downloadFile(roles.sourceRole, srcSite, srcServerRel);
    try {
      await SPm.uploadFile(roles.targetRole, dstSite, destFolder, name, blob, { overwrite: MoveRules.uploadOverwrite(cfg.overwrite, created), cancelToken: ctx.cancelToken });
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
    return undefined;
  }

  // A raw 403/423 on the target usually means a protected file, not a tool bug.
  function accessError(e, name) {
    if (e && (e.status === 403 || e.status === 423)) {
      const ne = new Error(`Access denied for ${name}: the target file is probably checked out, locked by a user, or a record / on hold. (${e.message})`);
      ne.status = e.status;
      return ne;
    }
    return e;
  }

  function fmtBytes(n) {
    if (!n) return '0 B';
    const u = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.min(u.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
    return `${(n / Math.pow(1024, i)).toFixed(i ? 1 : 0)} ${u[i]}`;
  }

  window.Transfer = { run, pause, resumeRun, cancel, isRunning, rolesInUse, fmtBytes };
})();
