(function () {
  'use strict';

  const cfg = window.MICROSOFT_CONFIG || {};
  const auth = window.MicrosoftAuth;
  const rootCache = { resolved: null, appRoot: null };
  const monthNames = window.TUKIN_CONFIG?.MONTHS || [
    'Januari','Februari','Maret','April','Mei','Juni',
    'Juli','Agustus','September','Oktober','November','Desember'
  ];

  function uid(prefix) {
    if (window.crypto?.randomUUID) return `${prefix}-${crypto.randomUUID()}`;
    return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function isFileLike(value) {
    return (typeof File !== 'undefined' && value instanceof File) ||
      (typeof Blob !== 'undefined' && value instanceof Blob);
  }

  function safeName(value, fallback) {
    let result = String(value || fallback || 'item')
      .replace(/[~#%&*{}\\:<>?/+|"\x00-\x1F]/g, '-')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/[. ]+$/g, '');
    if (!result) result = fallback || 'item';
    return result.slice(0, 120);
  }

  function extension(name) {
    const match = String(name || '').match(/(\.[A-Za-z0-9]{1,8})$/);
    return match ? match[1].toLowerCase() : '';
  }

  function base64Url(value) {
    const bytes = new TextEncoder().encode(String(value));
    let binary = '';
    bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
    return `u!${btoa(binary).replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-')}`;
  }

  function graphUrl(path) {
    const base = String(cfg.graphBaseUrl || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
    return `${base}${path.startsWith('/') ? path : `/${path}`}`;
  }

  async function graphFetch(path, options) {
    const opts = options || {};
    const token = await auth.acquireToken(Boolean(opts.interactive));
    const headers = new Headers(opts.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    if (opts.json !== undefined) headers.set('Content-Type', 'application/json');

    const response = await fetch(graphUrl(path), {
      method: opts.method || 'GET',
      headers,
      body: opts.json !== undefined ? JSON.stringify(opts.json) : opts.body,
      redirect: opts.redirect || 'follow'
    });

    if (!response.ok) {
      let detail = '';
      try {
        const payload = await response.json();
        detail = payload?.error?.message || payload?.error?.code || JSON.stringify(payload);
      } catch (_) {
        detail = await response.text().catch(() => '');
      }
      const error = new Error(`Microsoft Graph ${response.status}: ${detail || response.statusText}`);
      error.status = response.status;
      throw error;
    }
    return response;
  }

  async function graphJson(path, options) {
    const response = await graphFetch(path, options);
    if (response.status === 204) return null;
    return response.json();
  }

  async function ensureReady(interactive) {
    if (!auth?.isConfigured?.()) throw new Error('Konfigurasi Microsoft/SharePoint belum lengkap.');
    await auth.initialize();
    if (!auth.getStatus().connected) {
      if (!interactive) throw new Error('SharePoint belum terhubung.');
      await auth.connect();
    }
    await resolveAppRoot(Boolean(interactive));
    return true;
  }

  async function resolveSharedRoot(interactive) {
    if (rootCache.resolved) return rootCache.resolved;
    if (!cfg.shareUrl) throw new Error('shareUrl SharePoint belum diisi.');
    const encoded = base64Url(cfg.shareUrl);
    const item = await graphJson(`/shares/${encoded}/driveItem?$select=id,name,webUrl,parentReference,folder`, {
      interactive,
      headers: { Prefer: 'redeemSharingLink' }
    });
    const driveId = item?.parentReference?.driveId;
    if (!item?.id || !driveId || !item?.folder) {
      throw new Error('Tautan SharePoint tidak mengarah ke folder yang dapat ditulis.');
    }
    rootCache.resolved = {
      driveId,
      itemId: item.id,
      name: item.name,
      webUrl: item.webUrl || '',
      provider: 'sharepoint'
    };
    return rootCache.resolved;
  }

  async function listChildren(driveId, parentItemId, interactive) {
    const rows = [];
    let path = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentItemId)}/children?$select=id,name,webUrl,parentReference,folder,file,eTag,cTag,size,lastModifiedDateTime&$top=200`;
    while (path) {
      const page = await graphJson(path, { interactive });
      rows.push(...(page?.value || []));
      const next = page?.['@odata.nextLink'];
      path = next ? next.replace(String(cfg.graphBaseUrl || 'https://graph.microsoft.com/v1.0').replace(/\/$/, ''), '') : '';
    }
    return rows;
  }

  async function findChildFolder(driveId, parentItemId, name, interactive) {
    const rows = await listChildren(driveId, parentItemId, interactive);
    return rows.find((item) => item.folder && String(item.name).toLowerCase() === String(name).toLowerCase()) || null;
  }

  async function ensureFolder(driveId, parentItemId, rawName, interactive) {
    const name = safeName(rawName, 'Folder');
    const existing = await findChildFolder(driveId, parentItemId, name, interactive);
    if (existing) return existing;
    try {
      return await graphJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentItemId)}/children`, {
        method: 'POST',
        interactive,
        json: {
          name,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'fail'
        }
      });
    } catch (error) {
      if (error.status !== 409) throw error;
      const afterRace = await findChildFolder(driveId, parentItemId, name, interactive);
      if (afterRace) return afterRace;
      throw error;
    }
  }

  async function resolveAppRoot(interactive) {
    if (rootCache.appRoot) return rootCache.appRoot;
    const shared = await resolveSharedRoot(interactive);
    const appRoot = cfg.rootFolderName
      ? await ensureFolder(shared.driveId, shared.itemId, cfg.rootFolderName, interactive)
      : { id: shared.itemId, name: shared.name, webUrl: shared.webUrl, parentReference: { driveId: shared.driveId } };
    rootCache.appRoot = {
      driveId: shared.driveId,
      itemId: appRoot.id,
      name: appRoot.name,
      webUrl: appRoot.webUrl || '',
      provider: 'sharepoint'
    };
    return rootCache.appRoot;
  }

  function periodFolderName(period) {
    const month = String(Number(period?.month || 0)).padStart(2, '0');
    const label = monthNames[Number(period?.month || 1) - 1] || month;
    return `${period?.year}-${month}_${label}`;
  }

  function runFolderName(runId) {
    const short = String(runId || uid('run')).replace(/^tukin-/, '').replace(/[^A-Za-z0-9]/g, '').slice(0, 16);
    return `Run_${short || Date.now()}`;
  }

  function employeeFolderName(employee) {
    const nip = String(employee?.nip || '').replace(/\D/g, '') || 'TANPA-NIP';
    return safeName(`${nip}_${employee?.name || 'Pegawai'}`, nip);
  }

  async function ensureRunStructure(run, interactive) {
    await ensureReady(interactive);
    const appRoot = await resolveAppRoot(interactive);
    const driveId = run?.sharePoint?.driveId || appRoot.driveId;

    if (run?.sharePoint?.runFolderItemId && driveId) {
      return {
        driveId,
        runFolder: {
          id: run.sharePoint.runFolderItemId,
          name: run.sharePoint.runFolderName || runFolderName(run.id),
          webUrl: run.sharePoint.runFolderWebUrl || ''
        }
      };
    }

    const yearFolder = await ensureFolder(appRoot.driveId, appRoot.itemId, String(run.period.year), interactive);
    const periodFolder = await ensureFolder(appRoot.driveId, yearFolder.id, periodFolderName(run.period), interactive);
    const runsFolder = await ensureFolder(appRoot.driveId, periodFolder.id, 'Runs', interactive);
    const runFolder = await ensureFolder(appRoot.driveId, runsFolder.id, runFolderName(run.id), interactive);
    return { driveId: appRoot.driveId, runFolder };
  }

  function normalizeAttachment(item, kind) {
    if (isFileLike(item)) {
      return {
        id: uid(kind || 'file'),
        name: item.name || `${kind || 'file'}`,
        size: Number(item.size || 0),
        type: item.type || '',
        lastModified: Number(item.lastModified || 0),
        file: item,
        provider: 'local',
        localOnly: true
      };
    }
    const result = { ...(item || {}) };
    if (!result.id) result.id = uid(kind || 'file');
    if (!result.name && result.file?.name) result.name = result.file.name;
    if (!result.size && result.file?.size) result.size = Number(result.file.size || 0);
    if (!result.type && result.file?.type) result.type = result.file.type || '';
    if (!result.lastModified && result.file?.lastModified) result.lastModified = Number(result.file.lastModified || 0);
    if (!('file' in result)) result.file = null;
    if (result.itemId && result.driveId) {
      result.provider = 'sharepoint';
      result.localOnly = false;
    }
    return result;
  }

  function normalizeEmployeeAttachments(employee) {
    employee.sourceFiles = (employee.sourceFiles || []).map((item) => normalizeAttachment(item, 'absensi'));
    Object.values(employee.records || {}).forEach((record) => {
      record.evidence = (record.evidence || []).map((item) => normalizeAttachment(item, 'bukti'));
    });
    return employee;
  }

  function serializableAttachment(item) {
    const attachment = normalizeAttachment(item, 'file');
    return {
      id: attachment.id || '',
      name: attachment.name || '',
      size: Number(attachment.size || 0),
      type: attachment.type || '',
      lastModified: Number(attachment.lastModified || 0),
      provider: attachment.itemId && attachment.driveId ? 'sharepoint' : (attachment.provider || 'local'),
      driveId: attachment.driveId || '',
      itemId: attachment.itemId || '',
      webUrl: attachment.webUrl || '',
      eTag: attachment.eTag || '',
      cTag: attachment.cTag || '',
      remoteName: attachment.remoteName || '',
      remotePath: attachment.remotePath || '',
      sourceSize: Number(attachment.sourceSize || attachment.file?.size || attachment.size || 0),
      sourceLastModified: Number(attachment.sourceLastModified || attachment.file?.lastModified || attachment.lastModified || 0),
      uploadedAt: attachment.uploadedAt || '',
      uploadedBy: attachment.uploadedBy || '',
      localOnly: !(attachment.itemId && attachment.driveId)
    };
  }

  function needsUpload(attachment) {
    if (!attachment?.file) return false;
    if (!attachment.itemId || !attachment.driveId) return true;
    const size = Number(attachment.file.size || 0);
    const modified = Number(attachment.file.lastModified || 0);
    return size !== Number(attachment.sourceSize || attachment.size || 0) ||
      (modified && modified !== Number(attachment.sourceLastModified || attachment.lastModified || 0));
  }

  async function uploadBlob(driveId, parentItemId, remoteName, blob, interactive) {
    if (!blob) throw new Error('Isi file tidak tersedia untuk upload.');
    const max = Number(cfg.maxSimpleUploadBytes || 250 * 1024 * 1024);
    if (Number(blob.size || 0) > max) throw new Error(`File ${remoteName} melebihi batas upload ${Math.round(max / 1024 / 1024)} MB.`);
    const encodedName = encodeURIComponent(safeName(remoteName, 'file'));
    const item = await graphJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentItemId)}:/${encodedName}:/content`, {
      method: 'PUT',
      interactive,
      headers: { 'Content-Type': blob.type || 'application/octet-stream' },
      body: blob
    });
    return item;
  }

  function applyRemoteMetadata(attachment, remote, remoteName, remotePath) {
    const account = auth.getAccount?.();
    attachment.provider = 'sharepoint';
    attachment.driveId = remote?.parentReference?.driveId || attachment.driveId || '';
    attachment.itemId = remote?.id || attachment.itemId || '';
    attachment.webUrl = remote?.webUrl || attachment.webUrl || '';
    attachment.eTag = remote?.eTag || '';
    attachment.cTag = remote?.cTag || '';
    attachment.remoteName = remoteName || remote?.name || attachment.remoteName || '';
    attachment.remotePath = remotePath || attachment.remotePath || '';
    attachment.size = Number(attachment.file?.size || remote?.size || attachment.size || 0);
    attachment.sourceSize = Number(attachment.file?.size || attachment.sourceSize || attachment.size || 0);
    attachment.sourceLastModified = Number(attachment.file?.lastModified || attachment.sourceLastModified || attachment.lastModified || 0);
    attachment.uploadedAt = new Date().toISOString();
    attachment.uploadedBy = account?.username || account?.name || '';
    attachment.localOnly = false;
    return attachment;
  }

  async function syncRunFiles(run, interactive) {
    const structure = await ensureRunStructure(run, interactive);
    const driveId = structure.driveId;
    const runFolder = structure.runFolder;
    const employeeRoot = await ensureFolder(driveId, runFolder.id, 'Pegawai', interactive);
    const recapFolder = await ensureFolder(driveId, runFolder.id, 'Rekap', interactive);
    const uploaded = [];

    for (const employee of run.employees || []) {
      normalizeEmployeeAttachments(employee);
      const employeeFolder = await ensureFolder(driveId, employeeRoot.id, employeeFolderName(employee), interactive);
      const attendanceFolder = await ensureFolder(driveId, employeeFolder.id, 'Absensi', interactive);
      const evidenceFolder = await ensureFolder(driveId, employeeFolder.id, 'Bukti_Dukung', interactive);
      const nip = String(employee.nip || 'TANPA-NIP').replace(/\D/g, '') || 'TANPA-NIP';
      const safeEmployee = safeName(employee.name || 'Pegawai', 'Pegawai').replace(/\s+/g, '_');

      for (let index = 0; index < employee.sourceFiles.length; index += 1) {
        const attachment = employee.sourceFiles[index];
        if (!needsUpload(attachment)) continue;
        const ext = extension(attachment.name) || '.xlsx';
        const suffix = employee.sourceFiles.length > 1 ? `_${index + 1}` : '';
        const remoteName = safeName(`Absensi_${run.period.year}-${String(run.period.month).padStart(2, '0')}_${nip}_${safeEmployee}${suffix}${ext}`, `Absensi${ext}`);
        const remote = await uploadBlob(driveId, attendanceFolder.id, remoteName, attachment.file, interactive);
        applyRemoteMetadata(attachment, remote, remoteName, `Pegawai/${employeeFolder.name}/Absensi/${remoteName}`);
        uploaded.push(attachment);
      }

      for (const record of Object.values(employee.records || {})) {
        for (const attachment of record.evidence || []) {
          if (!needsUpload(attachment)) continue;
          const shortId = String(attachment.id || uid('bukti')).replace(/[^A-Za-z0-9]/g, '').slice(-8) || Date.now();
          const date = record.key || (record.date instanceof Date ? window.TukinRules?.dateKey?.(record.date) : '') || 'tanggal';
          const remoteName = safeName(`${date}_${shortId}_${attachment.name || attachment.file?.name || 'Bukti'}`, `Bukti_${shortId}`);
          const remote = await uploadBlob(driveId, evidenceFolder.id, remoteName, attachment.file, interactive);
          applyRemoteMetadata(attachment, remote, remoteName, `Pegawai/${employeeFolder.name}/Bukti_Dukung/${remoteName}`);
          uploaded.push(attachment);
        }
      }
    }

    const sharePoint = {
      provider: 'sharepoint',
      driveId,
      runFolderItemId: runFolder.id,
      runFolderName: runFolder.name,
      runFolderWebUrl: runFolder.webUrl || '',
      recapFolderItemId: recapFolder.id,
      rootShareUrl: cfg.shareUrl,
      rootFolderName: cfg.rootFolderName || '',
      syncedAt: new Date().toISOString()
    };
    run.sharePoint = sharePoint;
    return { sharePoint, uploadedCount: uploaded.length };
  }

  async function uploadRecap(run, recap, interactive) {
    if (!recap?.blob) return null;
    const structure = await ensureRunStructure(run, interactive);
    const driveId = structure.driveId;
    let recapFolderId = run?.sharePoint?.recapFolderItemId || '';
    if (!recapFolderId) {
      const folder = await ensureFolder(driveId, structure.runFolder.id, 'Rekap', interactive);
      recapFolderId = folder.id;
    }
    const remoteName = safeName(recap.name || 'Rekap_Tukin.xlsx', 'Rekap_Tukin.xlsx');
    const remote = await uploadBlob(driveId, recapFolderId, remoteName, recap.blob, interactive);
    const meta = applyRemoteMetadata(normalizeAttachment({
      id: 'rekap',
      name: recap.name,
      size: recap.blob.size,
      type: recap.blob.type,
      file: recap.blob
    }, 'rekap'), remote, remoteName, `Rekap/${remoteName}`);
    meta.file = null;
    run.sharePoint = { ...(run.sharePoint || {}), recap: serializableAttachment(meta), recapFolderItemId: recapFolderId };
    return run.sharePoint.recap;
  }

  async function downloadAttachment(item, interactive) {
    const attachment = normalizeAttachment(item, 'file');
    // Di browser, endpoint /content tidak dipakai langsung karena respons 302
    // dengan Authorization header dapat gagal pada CORS preflight. Ambil URL
    // download pre-authenticated dari metadata DriveItem, lalu fetch tanpa token.
    if (attachment.driveId && attachment.itemId) {
      const remote = await graphJson(
        `/drives/${encodeURIComponent(attachment.driveId)}/items/${encodeURIComponent(attachment.itemId)}?$select=id,name,size,webUrl,eTag,cTag,lastModifiedDateTime,file,@microsoft.graph.downloadUrl`,
        { interactive }
      );
      const downloadUrl = remote?.['@microsoft.graph.downloadUrl'];
      if (!downloadUrl) throw new Error(`URL download SharePoint tidak tersedia untuk ${attachment.name || attachment.itemId}.`);
      const response = await fetch(downloadUrl, { method: 'GET', credentials: 'omit' });
      if (!response.ok) throw new Error(`Download SharePoint gagal (${response.status}) untuk ${attachment.name || remote?.name || 'file'}.`);
      const blob = await response.blob();
      if (typeof File !== 'undefined') {
        return new File([blob], attachment.name || remote?.name || attachment.remoteName || 'file', {
          type: attachment.type || blob.type || 'application/octet-stream',
          lastModified: remote?.lastModifiedDateTime ? Date.parse(remote.lastModifiedDateTime) || Date.now() : Date.now()
        });
      }
      return blob;
    }
    return attachment.file || null;
  }

  async function refreshAttachmentMetadata(item, interactive) {
    const attachment = normalizeAttachment(item, 'file');
    if (!attachment.driveId || !attachment.itemId) return attachment;
    const remote = await graphJson(`/drives/${encodeURIComponent(attachment.driveId)}/items/${encodeURIComponent(attachment.itemId)}?$select=id,name,webUrl,parentReference,eTag,cTag,size,lastModifiedDateTime,file`, { interactive });
    applyRemoteMetadata(attachment, remote, attachment.remoteName || remote.name, attachment.remotePath || '');
    return attachment;
  }

  async function deleteAttachment(item, interactive) {
    const attachment = normalizeAttachment(item, 'file');
    if (!attachment.driveId || !attachment.itemId) return false;
    await graphFetch(`/drives/${encodeURIComponent(attachment.driveId)}/items/${encodeURIComponent(attachment.itemId)}`, {
      method: 'DELETE', interactive
    });
    return true;
  }

  async function deleteAttachments(items, interactive) {
    const failures = [];
    for (const item of items || []) {
      try { await deleteAttachment(item, interactive); }
      catch (error) { failures.push({ item, error }); }
    }
    return failures;
  }

  async function deleteRunFolder(sharePoint, interactive) {
    if (!sharePoint?.driveId || !sharePoint?.runFolderItemId) return false;
    await graphFetch(`/drives/${encodeURIComponent(sharePoint.driveId)}/items/${encodeURIComponent(sharePoint.runFolderItemId)}`, {
      method: 'DELETE', interactive
    });
    return true;
  }

  function hasRemote(item) {
    return Boolean(item?.driveId && item?.itemId);
  }

  function resetCache() {
    rootCache.resolved = null;
    rootCache.appRoot = null;
  }

  window.SharePointStorage = Object.freeze({
    ensureReady,
    resolveSharedRoot,
    resolveAppRoot,
    normalizeAttachment,
    normalizeEmployeeAttachments,
    serializableAttachment,
    syncRunFiles,
    uploadRecap,
    downloadAttachment,
    refreshAttachmentMetadata,
    deleteAttachment,
    deleteAttachments,
    deleteRunFolder,
    hasRemote,
    safeName,
    resetCache
  });
})();
