import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import './bcp.css';
import { createEmptyProject, refreshProjectCounts, validateProject } from './projectSchema';
import { deleteRecovery, getRecovery, listRecentProjects, saveRecentProject, saveRecovery } from './projectStorage';
import { downloadProjectArchive, readProjectArchive } from './projectPackage';
import { getDocumentView, searchDocuments, searchItems } from './registryEngine';
import { commitPartialImport, evaluateStagedImport } from './importEngine';
import { buildDuplicateReview, recordDuplicateDecision } from './duplicateReviewEngine';
import { applyStagingCorrection, getStagingValidationSummary } from './stagingEngine';
import { createMonthlyPeriod, changeMonthlyPeriodStatus, assignDocumentToPeriod, getMonthlyPeriodSummary } from './periodEngine';
import { buildExportWorkbook, downloadExportWorkbook, recordExportHistory, EXPORT_PRESETS, getExportFields, saveExportProfile } from './exportEngine';

function formatDate(value) {
  if (!value) return 'Belum pernah disimpan';
  return new Date(value).toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
}

function formatBytes(value) {
  if (value < 1024) return value + ' B';
  if (value < 1024 * 1024) return (value / 1024).toFixed(0) + ' KB';
  return (value / 1024 / 1024).toFixed(1) + ' MB';
}

export default function BCP() {
  const [project, setProject] = useState(null);
  const [recent, setRecent] = useState([]);
  const [view, setView] = useState('projects');
  const [documentQuery, setDocumentQuery] = useState('');
  const [itemQuery, setItemQuery] = useState('');
  const [selectedDocumentId, setSelectedDocumentId] = useState(null);
  const [stagedFiles, setStagedFiles] = useState([]);
  const [selectedDocumentIds, setSelectedDocumentIds] = useState([]);
  const [duplicateFilter, setDuplicateFilter] = useState('ALL');
  const [exportPreset, setExportPreset] = useState('DOC_REGISTRY');
  const [exportVersion, setExportVersion] = useState('CURRENT');
  const [exportEntityType, setExportEntityType] = useState('DOCUMENT');
  const [exportColumns, setExportColumns] = useState(['nomor_aju','nomor_daftar','document_type']);
  const [exportProfileName, setExportProfileName] = useState('');
  const [exportFilterField, setExportFilterField] = useState('');
  const [exportFilterValue, setExportFilterValue] = useState('');
  const [exportGroupBy, setExportGroupBy] = useState('');
  const [exportSortBy, setExportSortBy] = useState('');
  const [exportSortDirection, setExportSortDirection] = useState('ASC');
  const [exportLabelOverrides, setExportLabelOverrides] = useState('');
  const [directionOverrides, setDirectionOverrides] = useState({});
  const [stageBusy, setStageBusy] = useState(false);
  const [uploadDrag, setUploadDrag] = useState(false);
  const uploadInput = useRef(null);
  const [projectName, setProjectName] = useState('BCP PROJECT 2026');
  const [companyName, setCompanyName] = useState('PT Chinli Plastic Materials Indonesia');
  const [showCreate, setShowCreate] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [recoveryState, setRecoveryState] = useState('idle');
  const [reminder, setReminder] = useState(false);
  const fileInput = useRef(null);
  const changedAt = useRef(Date.now());
  const reminderShown = useRef(false);

  const isDirty = useMemo(() => {
    if (!project) return false;
    if (!project.manifest.last_saved_at) return true;
    return new Date(project.manifest.updated_at).getTime() > new Date(project.manifest.last_saved_at).getTime();
  }, [project]);

  const refreshRecent = useCallback(async () => {
    try { setRecent(await listRecentProjects()); }
    catch (e) { setError(e.message || 'Tidak dapat membaca daftar proyek lokal.'); }
  }, []);

  useEffect(() => { refreshRecent(); }, [refreshRecent]);

  useEffect(() => {
    if (!project) return undefined;
    const timer = window.setTimeout(async () => {
      try {
        await saveRecovery(project);
        await saveRecentProject(project);
        setRecoveryState('saved');
        refreshRecent();
      } catch (e) {
        setRecoveryState('error');
        setError('Auto-Recovery gagal: ' + (e.message || 'penyimpanan lokal tidak tersedia.'));
      }
    }, 450);
    return () => window.clearTimeout(timer);
  }, [project, refreshRecent]);

  useEffect(() => {
    if (!project || !isDirty) { setReminder(false); reminderShown.current = false; return undefined; }
    changedAt.current = Date.now();
    reminderShown.current = false;
    const timer = window.setInterval(() => {
      const minutes = Number(project.settings?.project?.reminder_minutes || 5);
      if (!reminderShown.current && Date.now() - changedAt.current >= minutes * 60 * 1000) {
        reminderShown.current = true;
        setReminder(true);
      }
    }, 30000);
    return () => window.clearInterval(timer);
  }, [project, isDirty]);

  const createProject = () => {
    if (!projectName.trim()) { setError('Nama proyek wajib diisi.'); return; }
    const next = createEmptyProject({ projectName, companyName });
    next.data.audit_log.push({
      event_id: crypto.randomUUID(), event_type: 'PROJECT_CREATED',
      description: 'Proyek BCP dibuat.', created_at: new Date().toISOString(),
      operator_id: next.settings.operator.operator_id
    });
    setProject(next);
    setView('dashboard');
    setShowCreate(false);
    setMessage('Proyek baru dibuat. Simpan file .bcp ke perangkat untuk penyimpanan permanen.');
    setError('');
  };

  const openArchive = async file => {
    if (!file) return;
    setBusy(true); setError(''); setMessage('');
    try {
      const loaded = await readProjectArchive(file);
      const recovery = await getRecovery(loaded.manifest.project_id);
      let selected = loaded;
      if (recovery?.project && new Date(recovery.project.manifest.updated_at) > new Date(loaded.manifest.updated_at)) {
        const restore = window.confirm('Ditemukan Auto-Recovery yang lebih baru dari file proyek. Pulihkan versi Auto-Recovery?');
        if (restore) {
          selected = recovery.project;
          setMessage('Versi Auto-Recovery berhasil dipulihkan. Simpan kembali sebagai file .bcp.');
        }
      }
      setProject(selected);
      setView('dashboard');
      await saveRecentProject(selected);
      await refreshRecent();
      setRecoveryState('saved');
    } catch (e) {
      setError(e.message || 'File proyek tidak dapat dibuka.');
    } finally {
      setBusy(false);
      if (fileInput.current) fileInput.current.value = '';
    }
  };

  const openRecent = async row => {
    setBusy(true); setError(''); setMessage('');
    try {
      const recovery = await getRecovery(row.project_id);
      if (!recovery?.project) {
        setMessage('Belum ada snapshot Auto-Recovery untuk proyek ini. Pilih file .bcp yang tersimpan di perangkat.');
        return;
      }
      setProject(recovery.project);
      setView('dashboard');
      setMessage('Proyek dibuka dari Auto-Recovery lokal. Simpan sebagai file .bcp untuk penyimpanan permanen.');
    } catch (e) { setError(e.message || 'Gagal membuka proyek terakhir.'); }
    finally { setBusy(false); }
  };

  const saveProject = async () => {
    if (!project) return;
    setBusy(true); setError(''); setMessage('');
    try {
      const savedManifest = await downloadProjectArchive(project);
      const next = {
        ...project,
        manifest: { ...project.manifest, last_saved_at: savedManifest.last_saved_at }
      };
      setProject(next);
      setMessage('File proyek .bcp telah dibuat. Pastikan file hasil unduhan tersimpan di lokasi yang kamu inginkan.');
      setReminder(false);
    } catch (e) { setError(e.message || 'Gagal membuat file proyek.'); }
    finally { setBusy(false); }
  };

  const closeProject = () => {
    if (isDirty && !window.confirm('Proyek memiliki perubahan setelah penyimpanan terakhir. Tutup proyek dan andalkan Auto-Recovery lokal?')) return;
    setProject(null);
    setView('projects');
    setMessage('Proyek ditutup. Pastikan file .bcp sudah disimpan jika ingin menyimpan permanen.');
    refreshRecent();
  };

  const addUploadFiles = async fileList => {
    if (!project) return;
    const files = [...(fileList || [])].filter(file => /\.(xlsx|xls|xlsm)$/i.test(file.name));
    if (!files.length) { setError('Pilih file Excel CEISA (.xlsx, .xls, atau .xlsm).'); return; }
    setStageBusy(true); setError(''); setMessage('Membaca dan mengidentifikasi file CEISA...');
    const parsed = [];
    const { parseBC27File } = await import('./ceisaParser.js');
    for (const file of files) {
      try {
        parsed.push(await parseBC27File(file, project.manifest.company_name));
      } catch (e) {
        parsed.push({ file_name: file.name, parse_error: e.message || 'File gagal diproses.' });
      }
    }
    setStagedFiles(current => [...current, ...parsed]);
    const readyIds = parsed.flatMap(file => (file.documents || []).filter(row => row.validation?.status !== 'ERROR' && row.document.document_type !== 'BC27_UNKNOWN').map(row => row.document.document_id));
    setSelectedDocumentIds(current => [...new Set([...current, ...readyIds])]);
    setMessage(`${parsed.length} file selesai diperiksa. Tinjau hasil klasifikasi dan duplikasi sebelum impor.`);
    setStageBusy(false);
    if (uploadInput.current) uploadInput.current.value = '';
  };

  const confirmPartialImport = () => {
    if (!project || !selectedDocumentIds.length) return;
    setStageBusy(true); setError('');
    try {
      const result = commitPartialImport(project, stagedFiles, selectedDocumentIds, directionOverrides);
      setProject(result.project);
      const imported = new Set(result.resolved_document_ids || result.imported_document_ids);
      setStagedFiles(current => current.map(file => file.parse_error ? file : ({ ...file, documents: file.documents.filter(entry => !imported.has(entry.document.document_id)) })).filter(file => file.parse_error || file.documents.length));
      setSelectedDocumentIds(current => current.filter(id => !imported.has(id)));
      setDirectionOverrides(current => Object.fromEntries(Object.entries(current).filter(([id]) => !imported.has(id))));
      setMessage(`Partial Import selesai: ${result.summary.imported_documents} dokumen baru diimpor, ${result.summary.linked_existing_documents || 0} sumber ditautkan ke dokumen existing, dan ${result.summary.held_documents} record masih ditahan.`);
    } catch (e) { setError(e.message || 'Import gagal.'); }
    finally { setStageBusy(false); }
  };

  const createPeriodFromPrompt = () => {
    const year = window.prompt('Tahun periode (contoh: 2026):', String(new Date().getFullYear()));
    if (year === null) return;
    const month = window.prompt('Bulan periode (1–12):', String(new Date().getMonth() + 1));
    if (month === null) return;
    try {
      setProject(current => createMonthlyPeriod(current, { year, month }));
      setMessage('Periode bulanan berhasil dibuat.');
      setError('');
    } catch (e) { setError(e.message || 'Gagal membuat periode.'); }
  };
  const updatePeriodStatus = (period, status) => {
    const reasonInput = window.prompt(
      status === 'CLOSED' || period.status === 'CLOSED'
        ? 'Alasan penutupan atau pembukaan kembali periode:'
        : 'Catatan perubahan status (opsional):'
    );
    if (reasonInput === null) return;
    const reason = reasonInput.trim();
    try {
      setProject(current => changeMonthlyPeriodStatus(current, period.period_id, status, reason));
      setMessage(`Status periode ${period.period_key} diperbarui menjadi ${status}.`);
      setError('');
    } catch (e) { setError(e.message || 'Gagal mengubah status periode.'); }
  };
  const changeDocumentPeriod = (documentId, periodId) => {
    const document = project?.data.documents.find(row => row.document_id === documentId);
    const reason = document?.period_id ? window.prompt('Alasan pemindahan dokumen antarperiode:') : '';
    if (reason === null) return;
    try {
      setProject(current => assignDocumentToPeriod(current, documentId, periodId, reason));
      setMessage('Periode dokumen berhasil diperbarui.');
      setError('');
    } catch (e) { setError(e.message || 'Gagal menetapkan periode dokumen.'); }
  };

  const exportNow = () => {
    if (!project) return;
    try {
      const label_overrides = Object.fromEntries(exportLabelOverrides.split(';').map(part => part.split('=').map(value => value.trim())).filter(pair => pair.length === 2 && pair[0] && pair[1]));
      const custom = exportPreset === 'CUSTOM' ? {
        entity_type: exportEntityType, columns: exportColumns,
        filters: exportFilterField && exportFilterValue ? { [exportFilterField]: exportFilterValue } : {},
        group_by: exportGroupBy, sort_by: exportSortBy, sort_direction: exportSortDirection, label_overrides
      } : null;
      const workbook = buildExportWorkbook(project, { preset: exportPreset, dataVersion: exportVersion, custom });
      const fileName = `BCP_${exportPreset}_${new Date().toISOString().slice(0,10)}.xlsx`;
      const bytes = downloadExportWorkbook(workbook, fileName);
      setProject(current => recordExportHistory(current, {
        preset: exportPreset, data_version: exportVersion, file_name: fileName,
        record_count: exportEntityType === 'ITEM' ? current.data.items.length : current.data.documents.length
      }));
      setMessage(`Export berhasil dibuat: ${fileName} (${formatBytes(bytes)}).`);
      setError('');
    } catch (e) { setError(e.message || 'Export gagal dibuat.'); }
  };
  const saveCurrentExportProfile = () => {
    try {
      setProject(current => saveExportProfile(current, {
        name: exportProfileName, entity_type: exportEntityType, columns: exportColumns,
        filters: exportFilterField && exportFilterValue ? { [exportFilterField]: exportFilterValue } : {},
        group_by: exportGroupBy, sort_by: exportSortBy, sort_direction: exportSortDirection,
        label_overrides: Object.fromEntries(exportLabelOverrides.split(';').map(part => part.split('=').map(value => value.trim())).filter(pair => pair.length === 2 && pair[0] && pair[1]))
      }));
      setMessage('Profil Custom Export berhasil disimpan dalam proyek.');
      setExportProfileName('');
      setError('');
    } catch (e) { setError(e.message || 'Profil export gagal disimpan.'); }
  };
  const loadExportProfile = profileId => {
    const profile = project?.settings?.export?.profiles?.find(row => row.profile_id === profileId);
    if (!profile) return;
    setExportEntityType(profile.entity_type);
    setExportColumns(profile.columns);
    const filterEntry = Object.entries(profile.filters || {})[0] || ['', ''];
    setExportFilterField(filterEntry[0]); setExportFilterValue(filterEntry[1]);
    setExportGroupBy(profile.group_by || ''); setExportSortBy(profile.sort_by || '');
    setExportSortDirection(profile.sort_direction || 'ASC');
    setExportLabelOverrides(Object.entries(profile.label_overrides || {}).map(([key,value]) => `${key}=${value}`).join('; '));
    setExportPreset('CUSTOM');
  };

  const updateOperator = value => {
    setProject(current => refreshProjectCounts({
      ...current,
      settings: { ...current.settings, operator: { ...current.settings.operator, display_name: value } }
    }));
  };

  const validationErrors = project ? validateProject(project) : [];
  const documentRows = project ? searchDocuments(project, documentQuery) : [];
  const itemRows = project ? searchItems(project, itemQuery) : [];
  const selectedDocument = project && selectedDocumentId ? getDocumentView(project, selectedDocumentId) : null;
  const stagedEvaluation = project ? evaluateStagedImport(project, stagedFiles) : [];
  const duplicateCandidates = project ? buildDuplicateReview(project, stagedFiles) : [];
  const stagingValidation = getStagingValidationSummary(stagedFiles);
  const correctStagedValue = (entry, entityType, entityId, fieldPath, label, currentValue) => {
    const value = window.prompt(`Nilai baru untuk ${label}:`, currentValue == null ? '' : String(currentValue));
    if (value === null) return;
    const reason = window.prompt('Alasan koreksi staging:');
    if (!reason || !reason.trim()) return;
    try {
      setStagedFiles(current => applyStagingCorrection(current, {
        document_id: entry.document.document_id, entity_type: entityType, entity_id: entityId,
        field_path: fieldPath, new_value: value, reason,
        operator_id: project?.settings?.operator?.operator_id || 'LOCAL-USER-001'
      }));
      setMessage('Koreksi staging disimpan. Nilai asli CEISA dan file sumber tetap tidak berubah.');
      setError('');
    } catch (e) { setError(e.message || 'Koreksi staging gagal.'); }
  };
  const decideDuplicate = (candidate, decision) => {
    const reason = window.prompt('Masukkan alasan keputusan duplikasi:');
    if (!reason || !reason.trim()) return;
    try {
      setProject(current => recordDuplicateDecision(current, candidate, decision, reason));
      if (candidate.staged_document_id && ['USE_EXISTING', 'USE_NEW', 'KEEP_BOTH'].includes(decision)) {
        setSelectedDocumentIds(current => current.includes(candidate.staged_document_id) ? current : [...current, candidate.staged_document_id]);
      } else if (candidate.staged_document_id) {
        setSelectedDocumentIds(current => current.filter(id => id !== candidate.staged_document_id));
      }
      setMessage('Keputusan Duplicate Review dicatat dalam Audit Trail. Periksa kembali Import Preview sebelum commit.');
      setError('');
    } catch (e) { setError(e.message || 'Keputusan duplikasi gagal disimpan.'); }
  };
  const go = next => { setView(next); setError(''); };

  return (
    <main className="bcp-app">
      <header className="bcp-header">
        <div className="bcp-brand-mark">BCP</div>
        <div className="bcp-heading">
          <div className="bcp-eyebrow">REZA SYSTEM · LOCAL PROJECT ENGINE</div>
          <h1>Buku Catatan Pemasukan &amp; Pengeluaran</h1>
          <p>BCP · Buku Catatan Pemasukan dan Pengeluaran</p>
        </div>
        <div className="bcp-header-status"><span className="bcp-status-dot"/><span>LOCAL-FIRST</span></div>
      </header>

      {project && (
        <section className="bcp-project-strip">
          <div className="bcp-project-ident">
            <span className="bcp-label">PROYEK AKTIF</span>
            <strong>{project.manifest.project_name}</strong>
            <small>{project.manifest.company_name || 'Nama perusahaan belum diisi'}</small>
          </div>
          <div className="bcp-save-state">
            <span className={isDirty ? 'bcp-state-dirty' : 'bcp-state-saved'}>{isDirty ? 'BELUM DISIMPAN KE FILE' : 'FILE PROYEK TERSIMPAN'}</span>
            <small>Auto-Recovery: {recoveryState === 'saved' ? 'Aktif' : recoveryState === 'error' ? 'Bermasalah' : 'Menyiapkan...'}</small>
          </div>
          <div className="bcp-strip-actions">
            <button className="bcp-button bcp-button-primary" onClick={saveProject} disabled={busy}>Simpan .bcp</button>
            <button className="bcp-button bcp-button-quiet" onClick={closeProject}>Tutup Proyek</button>
          </div>
        </section>
      )}

      <nav className="bcp-nav" aria-label="Navigasi BCP">
        <button className={view === 'projects' ? 'active' : ''} onClick={() => go('projects')}>Project Management</button>
        {project && <button className={view === 'dashboard' ? 'active' : ''} onClick={() => go('dashboard')}>Dashboard</button>}
        {project && <button className={view === 'upload' ? 'active' : ''} onClick={() => go('upload')}>Upload CEISA</button>}
        {project && <button className={view === 'documents' ? 'active' : ''} onClick={() => { setSelectedDocumentId(null); go('documents'); }}>Document Registry</button>}
        {project && <button className={view === 'items' ? 'active' : ''} onClick={() => go('items')}>Item Registry</button>}
        {project && <button className={view === 'duplicates' ? 'active' : ''} onClick={() => go('duplicates')}>Duplicate Review {duplicateCandidates.length ? `(${duplicateCandidates.length})` : ''}</button>}
        {project && <button className={view === 'validation' ? 'active' : ''} onClick={() => go('validation')}>Validation Center</button>}
        {project && <button className={view === 'periods' ? 'active' : ''} onClick={() => go('periods')}>Monthly Period</button>}
        {project && <button className={view === 'exports' ? 'active' : ''} onClick={() => go('exports')}>Export Center</button>}
        {project && <button className={view === 'settings' ? 'active' : ''} onClick={() => go('settings')}>Pengaturan Proyek</button>}
        {project && <button className={view === 'diagnostics' ? 'active' : ''} onClick={() => go('diagnostics')}>System Diagnostics</button>}
      </nav>

      {message && <div className="bcp-notice">{message}</div>}
      {error && <div className="bcp-error" role="alert">{error}</div>}
      {reminder && <div className="bcp-reminder"><div><strong>Pengingat penyimpanan</strong><span>Perubahan proyek belum disimpan ke file .bcp. Auto-Recovery bukan pengganti file proyek permanen.</span></div><button className="bcp-button bcp-button-primary" onClick={saveProject}>Simpan Sekarang</button><button className="bcp-dismiss" onClick={() => setReminder(false)} aria-label="Tutup pengingat">×</button></div>}

      {view === 'projects' && (
        <section className="bcp-content">
          <div className="bcp-section-heading">
            <div><span className="bcp-eyebrow">PROJECT WORKSPACE</span><h2>Project Management</h2><p>Buat proyek baru, buka file .bcp, atau lanjutkan pekerjaan dari Auto-Recovery.</p></div>
            <div className="bcp-heading-actions">
              <button className="bcp-button bcp-button-secondary" onClick={() => fileInput.current?.click()} disabled={busy}>Buka File .bcp</button>
              <button className="bcp-button bcp-button-primary" onClick={() => setShowCreate(true)}>＋ Proyek Baru</button>
              <input ref={fileInput} type="file" accept=".bcp,application/zip" hidden onChange={e => openArchive(e.target.files?.[0])}/>
            </div>
          </div>

          <div className="bcp-project-welcome">
            <div className="bcp-welcome-icon">▤</div>
            <div><strong>Data BCP berada di perangkat kamu</strong><p>File .bcp adalah penyimpanan permanen proyek. Auto-Recovery menggunakan IndexedDB browser untuk membantu memulihkan perubahan yang belum disimpan.</p></div>
          </div>

          <div className="bcp-section-title"><h3>Proyek Terakhir</h3><button className="bcp-text-button" onClick={refreshRecent}>Muat ulang</button></div>
          {recent.length === 0 ? (
            <div className="bcp-empty"><div>▧</div><strong>Belum ada proyek terakhir</strong><span>Buat proyek baru atau buka file .bcp yang sudah kamu miliki.</span></div>
          ) : (
            <div className="bcp-recent-list">
              {recent.map(row => (
                <article className="bcp-recent-card" key={row.project_id}>
                  <div className="bcp-file-icon">BCP</div>
                  <div className="bcp-recent-info"><strong>{row.project_name}</strong><span>{row.company_name || '—'}</span><small>Aktivitas terakhir: {formatDate(row.updated_at)}</small></div>
                  <button className="bcp-button bcp-button-secondary" onClick={() => openRecent(row)} disabled={busy}>Lanjutkan</button>
                </article>
              ))}
            </div>
          )}
          <div className="bcp-security-note"><strong>Catatan penyimpanan</strong><span>BCP tidak menggunakan backup otomatis, enkripsi, atau penyimpanan server. Jika file .bcp dan data Auto-Recovery browser hilang atau rusak, pemulihan mungkin tidak dapat dilakukan.</span></div>
        </section>
      )}

      {view === 'dashboard' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">PROJECT OVERVIEW</span><h2>Dashboard BCP</h2><p>Ringkasan proyek lokal, dokumen CEISA, barang, file sumber, dan aktivitas pencatatan.</p></div></div>
          <div className="bcp-metric-grid">
            <div><span>Dokumen</span><strong>{project.data.documents.length.toLocaleString('id-ID')}</strong><small>Document Registry</small></div>
            <div><span>Detail Barang</span><strong>{project.data.items.length.toLocaleString('id-ID')}</strong><small>Item Registry</small></div>
            <div><span>File Sumber</span><strong>{project.sources.source_registry.length.toLocaleString('id-ID')}</strong><small>Source Registry</small></div>
            <div><span>Riwayat Aktivitas</span><strong>{project.data.audit_log.length.toLocaleString('id-ID')}</strong><small>Audit Trail lokal</small></div>
          </div>
          <div className="bcp-next-stage">
            <span className="bcp-eyebrow">DEVELOPMENT ROADMAP</span>
            <h3>Fondasi proyek dan impor BC 2.7 sudah tersedia</h3>
            <p>Project Engine, Registry, Unified Upload, parser BC 2.7, Duplicate Review, dan Validation Center sedang dikembangkan secara bertahap.</p>
            <div className="bcp-progress"><span style={{ width: '50%' }}/></div>
            <small>M1–M4 · Foundation, Registry, Extraction, dan Import Control — development</small>
          </div>
          <div className="bcp-section-title"><h3>Informasi Proyek</h3></div>
          <div className="bcp-info-grid"><div><span>Project ID</span><code>{project.manifest.project_id}</code></div><div><span>Format</span><strong>{project.manifest.format} · {project.manifest.format_version}</strong></div><div><span>Schema Version</span><strong>{project.manifest.schema_version}</strong></div><div><span>Pembaruan</span><strong>{formatDate(project.manifest.updated_at)}</strong></div></div>
        </section>
      )}

      {view === 'upload' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">UNIFIED CEISA UPLOAD</span><h2>Upload BC 2.7</h2><p>Unggah Excel satu dokumen, workbook gabungan banyak dokumen, atau campuran keduanya dalam satu batch.</p></div></div>
          <div className={'bcp-upload-drop ' + (uploadDrag ? 'dragging' : '')}
            onDragOver={e => { e.preventDefault(); setUploadDrag(true); }}
            onDragLeave={() => setUploadDrag(false)}
            onDrop={e => { e.preventDefault(); setUploadDrag(false); addUploadFiles(e.dataTransfer.files); }}>
            <div className="bcp-upload-icon">⇧</div><strong>Tarik &amp; lepas file Excel CEISA di sini</strong><span>atau pilih beberapa file sekaligus</span>
            <button className="bcp-button bcp-button-secondary" onClick={() => uploadInput.current?.click()} disabled={stageBusy}>Pilih File Excel</button>
            <input ref={uploadInput} type="file" accept=".xlsx,.xls,.xlsm" multiple hidden onChange={e => addUploadFiles(e.target.files)}/>
            <small>BC 2.7 IN dan BC 2.7 OUT · Single Document dan Multi Document</small>
          </div>
          {stageBusy && <div className="bcp-notice">Memproses workbook di browser. Jangan tutup halaman ini.</div>}
          <div className="bcp-import-summary">
            <div><strong>{stagedFiles.length}</strong><span>File dalam staging</span></div>
            <div><strong>{stagedFiles.reduce((sum, file) => sum + (file.documents?.length || 0), 0)}</strong><span>Dokumen ditemukan</span></div>
            <div><strong>{stagedEvaluation.flatMap(file => file.documents || []).filter(row => row.status === 'READY').length}</strong><span>Siap diimpor</span></div>
            <div><strong>{stagedEvaluation.flatMap(file => file.documents || []).filter(row => ['ERROR','POSSIBLE_DUPLICATE','NEEDS_REVIEW'].includes(row.status)).length}</strong><span>Perlu pemeriksaan</span></div>
          </div>
          {stagedFiles.length > 0 && (
            <>
              <div className="bcp-section-title"><h3>Import Preview</h3><button className="bcp-text-button" onClick={() => { setStagedFiles([]); setSelectedDocumentIds([]); setDirectionOverrides({}); }}>Bersihkan staging</button></div>
              {stagedEvaluation.map((file, fileIndex) => (
                <article className="bcp-stage-file" key={file.source_id || file.file_name || fileIndex}>
                  <div className="bcp-stage-file-head"><div><strong>{file.file_name}</strong><small>{file.file_format || 'File gagal dibaca'} · {file.documents?.length || 0} dokumen</small></div><span className={'bcp-stage-badge ' + (file.status === 'STAGED' ? 'ready' : 'warning')}>{file.status}</span></div>
                  {file.reason && <div className="bcp-stage-reason">{file.reason}</div>}
                  {file.documents?.length > 0 && <div className="bcp-registry-table-wrap"><table className="bcp-registry-table"><thead><tr><th>Impor</th><th>Jenis BC</th><th>Nomor Aju</th><th>Nomor Daftar</th><th>Barang</th><th>Status</th><th>Arah transaksi</th></tr></thead><tbody>
                    {file.documents.map(row => {
                      const canSelect = ['READY','NEEDS_REVIEW','READY_WITH_DUPLICATE_DECISION','USE_EXISTING'].includes(row.status);
                      const override = directionOverrides[row.document_id] || '';
                      return <tr key={row.document_id}>
                        <td><input type="checkbox" checked={selectedDocumentIds.includes(row.document_id)} disabled={!canSelect || stageBusy} onChange={e => setSelectedDocumentIds(current => e.target.checked ? [...current, row.document_id] : current.filter(id => id !== row.document_id))}/></td>
                        <td>{row.document_type}</td><td>{row.aju_number || '—'}</td><td>{row.registration_number || '—'}</td><td>{row.item_count}</td>
                        <td><span className={'bcp-stage-badge ' + (row.status === 'READY' ? 'ready' : 'warning')}>{row.status}</span><small className="bcp-stage-reason">{row.reason}</small></td>
                        <td>{row.status === 'NEEDS_REVIEW' ? <select value={override} onChange={e => { const value = e.target.value; setDirectionOverrides(current => ({ ...current, [row.document_id]: value })); if (value) setSelectedDocumentIds(current => current.includes(row.document_id) ? current : [...current, row.document_id]); }}><option value="">Pilih arah</option><option value="IN">BC 2.7 IN</option><option value="OUT">BC 2.7 OUT</option></select> : row.document_type.replace('BC27_', '')}</td>
                      </tr>;
                    })}
                  </tbody></table></div>}
                  {file.issues?.length > 0 && <details className="bcp-stage-issues"><summary>Detail validasi file ({file.issues.length})</summary>{file.issues.map((issue, i) => <div key={issue.code + i}><b>{issue.severity}</b> · {issue.message}</div>)}</details>}
                </article>
              ))}
              <div className="bcp-import-actions"><span>{selectedDocumentIds.length} dokumen dipilih untuk Partial Import.</span><button className="bcp-button bcp-button-primary" onClick={confirmPartialImport} disabled={stageBusy || selectedDocumentIds.length === 0}>Konfirmasi Partial Import</button></div>
            </>
          )}
          <div className="bcp-security-note"><strong>Aturan impor</strong><span>Data sumber CEISA tidak diubah. Dokumen valid dan unik dapat diimpor terpisah; error, duplikasi, dan klasifikasi yang belum jelas tetap ditahan. File sumber yang sudah dipakai akan disertakan dalam proyek .bcp.</span></div>
        </section>
      )}

      {view === 'exports' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading">
            <div><span className="bcp-eyebrow">STANDARD &amp; CUSTOM EXPORT</span><h2>Export Center</h2><p>Buat workbook Excel dari Document Registry, Item Registry, periode, sumber, koreksi, dan Audit Trail. Export tidak mengubah Registry.</p></div>
          </div>
          <div className="bcp-export-form">
            <label>Standard Export Preset
              <select value={exportPreset} onChange={e => setExportPreset(e.target.value)}>
                {EXPORT_PRESETS.map(preset => <option key={preset} value={preset}>{preset}</option>)}
                <option value="CUSTOM">CUSTOM</option>
              </select>
            </label>
            <label>Data Version
              <select value={exportVersion} onChange={e => setExportVersion(e.target.value)}>
                <option value="CURRENT">Current (default)</option>
                <option value="ORIGINAL">Original</option>
                <option value="COMPARISON">Comparison Export</option>
              </select>
            </label>
            <label>Profil tersimpan
              <select defaultValue="" onChange={e => loadExportProfile(e.target.value)}>
                <option value="">Pilih profil Custom Export</option>
                {(project.settings.export.profiles || []).map(profile => <option key={profile.profile_id} value={profile.profile_id}>{profile.name}</option>)}
              </select>
            </label>
          </div>
          {exportPreset === 'CUSTOM' && (
            <div className="bcp-detail-panel">
              <div className="bcp-section-title"><h3>Custom Export</h3><span>{exportColumns.length} kolom dipilih</span></div>
              <label className="bcp-period-assignment">Jenis Registry
                <select value={exportEntityType} onChange={e => { setExportEntityType(e.target.value); setExportColumns([]); }}>
                  <option value="DOCUMENT">Document Registry</option><option value="ITEM">Item Registry</option>
                </select>
              </label>
              <div className="bcp-export-columns">
                {getExportFields(exportEntityType).map(field => (
                  <label key={field.key}><input type="checkbox" checked={exportColumns.includes(field.key)} onChange={e => setExportColumns(current => e.target.checked ? [...current, field.key] : current.filter(key => key !== field.key))}/><span>{field.label}</span></label>
                ))}
              </div>
              <div className="bcp-export-advanced">
                <label>Filter field<select value={exportFilterField} onChange={e => setExportFilterField(e.target.value)}><option value="">Tanpa filter</option>{getExportFields(exportEntityType).map(field => <option key={field.key} value={field.key}>{field.label}</option>)}</select></label>
                <label>Nilai filter<input value={exportFilterValue} onChange={e => setExportFilterValue(e.target.value)} placeholder="Cari nilai yang mengandung teks ini"/></label>
                <label>Group by<select value={exportGroupBy} onChange={e => setExportGroupBy(e.target.value)}><option value="">Tanpa grouping</option>{getExportFields(exportEntityType).filter(field => exportColumns.includes(field.key)).map(field => <option key={field.key} value={field.key}>{field.label}</option>)}</select></label>
                <label>Sort by<select value={exportSortBy} onChange={e => setExportSortBy(e.target.value)}><option value="">Urutan default</option>{getExportFields(exportEntityType).filter(field => exportColumns.includes(field.key)).map(field => <option key={field.key} value={field.key}>{field.label}</option>)}</select></label>
                <label>Arah sorting<select value={exportSortDirection} onChange={e => setExportSortDirection(e.target.value)}><option value="ASC">Ascending</option><option value="DESC">Descending</option></select></label>
                <label className="bcp-export-labels">Label kolom khusus<input value={exportLabelOverrides} onChange={e => setExportLabelOverrides(e.target.value)} placeholder="item_code=Kode Internal; quantity=Qty"/></label>
              </div>
              <div className="bcp-registry-toolbar">
                <input value={exportProfileName} onChange={e => setExportProfileName(e.target.value)} placeholder="Nama profil export"/>
                <button className="bcp-button bcp-button-secondary" onClick={saveCurrentExportProfile}>Simpan Profil</button>
              </div>
              <p className="bcp-muted">Profil menyimpan kolom, filter, grouping, sorting, dan label kolom. Filter menggunakan pencarian teks yang mengandung nilai.</p>
            </div>
          )}
          <div className="bcp-detail-panel">
            <div className="bcp-section-title"><h3>Preview Konfigurasi Export</h3><span>{exportVersion}</span></div>
            <div className="bcp-info-grid">
              <div><span>Preset</span><strong>{exportPreset}</strong></div>
              <div><span>Dokumen</span><strong>{project.data.documents.length.toLocaleString('id-ID')}</strong></div>
              <div><span>Barang</span><strong>{project.data.items.length.toLocaleString('id-ID')}</strong></div>
              <div><span>File sumber</span><strong>{project.sources.source_registry.length.toLocaleString('id-ID')}</strong></div>
            </div>
            {exportPreset === 'CUSTOM' && <p className="bcp-muted">Kolom: {getExportFields(exportEntityType).filter(field => exportColumns.includes(field.key)).map(field => field.label).join(' · ') || 'Belum ada kolom dipilih'}</p>}
            <div className="bcp-import-actions">
              <span>Workbook akan memuat README, VALIDATION_SUMMARY, dan EXPORT_METADATA.</span>
              <button className="bcp-button bcp-button-primary" onClick={exportNow} disabled={exportPreset === 'CUSTOM' && exportColumns.length === 0}>Export Excel (.xlsx)</button>
            </div>
          </div>
          <div className="bcp-section-title"><h3>Riwayat Export</h3><span>{project.data.export_history.length} export</span></div>
          {project.data.export_history.length === 0 ? <div className="bcp-empty"><div>⇩</div><strong>Belum ada riwayat export</strong><span>Riwayat akan dicatat setelah workbook berhasil dibuat.</span></div> : (
            <div className="bcp-registry-table-wrap"><table className="bcp-registry-table"><thead><tr><th>Waktu</th><th>Preset</th><th>Data Version</th><th>Nama File</th><th>Record</th></tr></thead><tbody>
              {[...project.data.export_history].reverse().map(row => <tr key={row.export_id}><td>{formatDate(row.exported_at)}</td><td>{row.preset}</td><td>{row.data_version}</td><td>{row.file_name}</td><td>{row.record_count}</td></tr>)}
            </tbody></table></div>
          )}
        </section>
      )}

      {view === 'periods' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading">
            <div><span className="bcp-eyebrow">MONTHLY PERIOD MANAGEMENT</span><h2>Manajemen Periode Bulanan</h2><p>Kelola periode OPEN, REVIEW, dan CLOSED. Penetapan periode dokumen dilakukan secara eksplisit; BCP belum menebak Reporting Date otomatis.</p></div>
            <button className="bcp-button bcp-button-primary" onClick={createPeriodFromPrompt}>＋ Buat Periode</button>
          </div>
          {project.data.periods.length === 0 ? (
            <div className="bcp-empty"><div>▦</div><strong>Belum ada periode</strong><span>Buat periode bulanan untuk mulai mengelompokkan dokumen BCP.</span></div>
          ) : [...project.data.periods].sort((a,b) => b.period_key.localeCompare(a.period_key)).map(period => {
            const summary = getMonthlyPeriodSummary(project, period.period_id);
            return (
              <article className="bcp-stage-file" key={period.period_id}>
                <div className="bcp-stage-file-head">
                  <div><strong>{period.label}</strong><small>{period.period_key} · Dibuat {formatDate(period.created_at)}</small></div>
                  <span className={'bcp-stage-badge ' + (period.status === 'OPEN' ? 'ready' : 'warning')}>{period.status}</span>
                </div>
                <div className="bcp-import-summary">
                  <div><strong>{summary.document_count}</strong><span>Dokumen</span></div>
                  <div><strong>{summary.item_count}</strong><span>Detail barang</span></div>
                  <div><strong>{summary.amended_count}</strong><span>Dokumen dikoreksi</span></div>
                  <div><strong>{summary.void_count + summary.cancelled_count}</strong><span>VOID / CANCELLED</span></div>
                </div>
                <div className="bcp-info-grid">
                  <div><span>Rincian jenis BC</span><strong>{Object.entries(summary.documents_by_type).map(([type,count]) => `${type}: ${count}`).join(' · ') || 'Belum ada dokumen'}</strong></div>
                  <div><span>Tanggal status terakhir</span><strong>{formatDate(period.updated_at)}</strong></div>
                </div>
                <div className="bcp-import-actions">
                  <span>Perubahan periode dicatat pada Audit Trail.</span>
                  <div className="bcp-heading-actions">
                    {period.status === 'OPEN' && <button className="bcp-button bcp-button-secondary" onClick={() => updatePeriodStatus(period, 'REVIEW')}>Mulai Review</button>}
                    {period.status === 'REVIEW' && <button className="bcp-button bcp-button-secondary" onClick={() => updatePeriodStatus(period, 'OPEN')}>Kembali OPEN</button>}
                    {period.status !== 'CLOSED' && <button className="bcp-button bcp-button-quiet" onClick={() => updatePeriodStatus(period, 'CLOSED')}>Tutup Periode</button>}
                    {period.status === 'CLOSED' && <button className="bcp-button bcp-button-primary" onClick={() => updatePeriodStatus(period, 'OPEN')}>Buka Kembali</button>}
                  </div>
                </div>
              </article>
            );
          })}
          <div className="bcp-security-note"><strong>Aturan Reporting Date</strong><span>BCP belum otomatis memilih Tanggal Aju, Tanggal Daftar, atau tanggal transaksi sebagai dasar periode. Penetapan otomatis per jenis BC akan diaktifkan setelah mapping tanggal CEISA disepakati dan diuji.</span></div>
        </section>
      )}

      {view === 'validation' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading">
            <div><span className="bcp-eyebrow">STAGING VALIDATION &amp; CORRECTION</span><h2>Validation Center</h2><p>Periksa error dan lakukan koreksi pada nilai hasil ekstraksi sebelum Partial Import. Nilai sumber asli tidak pernah ditimpa.</p></div>
          </div>
          {stagingValidation.length === 0 ? (
            <div className="bcp-empty"><div>✓</div><strong>Belum ada data staging</strong><span>Unggah file CEISA terlebih dahulu pada halaman Upload CEISA.</span></div>
          ) : stagingValidation.map((file, index) => (
            <article className="bcp-stage-file" key={file.source_id || file.file_name + index}>
              <div className="bcp-stage-file-head"><div><strong>{file.file_name}</strong><small>{file.documents.length} dokumen · {file.file_error ? 'File Error' : 'Sudah diekstrak'}</small></div><span className={'bcp-stage-badge ' + (file.file_error ? 'warning' : 'ready')}>{file.file_error ? 'FILE ERROR' : 'STAGING'}</span></div>
              {file.file_error && <div className="bcp-error">{file.file_error}</div>}
              {file.documents.map(summary => {
                const stagedEntry = stagedFiles.find(row => row.source?.source_id === file.source_id)?.documents?.find(row => row.document.document_id === summary.document_id);
                if (!stagedEntry) return null;
                return (
                  <div className="bcp-detail-panel" key={summary.document_id}>
                    <div className="bcp-section-title"><h3>{summary.document_type} · {summary.aju_number || 'Tanpa Nomor Aju'}</h3><span className={'bcp-stage-badge ' + (summary.status === 'READY' ? 'ready' : 'warning')}>{summary.status}</span></div>
                    <div className="bcp-info-grid">
                      <div><span>Nomor Daftar</span><strong>{stagedEntry.document.registration_number || '—'}</strong><button className="bcp-text-button" onClick={() => correctStagedValue(stagedEntry, 'DOCUMENT', stagedEntry.document.document_id, 'current_data.nomor_daftar', 'Nomor Daftar', stagedEntry.document.registration_number)}>Koreksi Nomor Daftar</button></div>
                      <div><span>Tanggal Daftar</span><strong>{stagedEntry.document.dates?.registration || '—'}</strong><button className="bcp-text-button" onClick={() => correctStagedValue(stagedEntry, 'DOCUMENT', stagedEntry.document.document_id, 'current_data.tanggal_daftar', 'Tanggal Daftar', stagedEntry.document.dates?.registration)}>Koreksi Tanggal Daftar</button></div>
                    </div>
                    {stagedEntry.items.map(item => (
                      <div className="bcp-registry-toolbar" key={item.item_id}>
                        <span>Seri {item.item_sequence} · {item.item_code || 'Tanpa Kode'}</span>
                        <span>{item.quantity ?? '—'} {item.unit || ''}</span>
                        <button className="bcp-text-button" onClick={() => correctStagedValue(stagedEntry, 'ITEM', item.item_id, 'current_data.description', 'Uraian Barang', item.description)}>Koreksi Uraian</button>
                        <button className="bcp-text-button" onClick={() => correctStagedValue(stagedEntry, 'ITEM', item.item_id, 'current_data.quantity', 'Jumlah Satuan', item.quantity)}>Koreksi Jumlah</button>
                      </div>
                    ))}
                    {summary.issues.length > 0 && <details className="bcp-stage-issues"><summary>Temuan validasi ({summary.issues.length})</summary>{summary.issues.map((issue, issueIndex) => <div key={issue.code + issueIndex}><b>{issue.severity}</b> · {issue.message}</div>)}</details>}
                    {summary.corrections.length > 0 && <div className="bcp-stage-issues"><b>{summary.corrections.length} koreksi staging</b>{summary.corrections.map(correction => <div key={correction.correction_id}>{correction.field_path}: {String(correction.old_value ?? '—')} → {String(correction.new_value ?? '—')} · {correction.reason}</div>)}</div>}
                  </div>
                );
              })}
            </article>
          ))}
          <div className="bcp-security-note"><strong>Batas koreksi staging</strong><span>Field sensitif seperti Nomor Aju, identitas sumber, dan Source ID tidak dapat diedit melalui Validation Center. Koreksi disimpan sebagai riwayat dan baru masuk ke Audit Trail proyek ketika dokumen berhasil diimpor.</span></div>
        </section>
      )}

      {view === 'duplicates' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading">
            <div><span className="bcp-eyebrow">DUPLICATE CONTROL CENTER</span><h2>Duplicate Review</h2><p>Bandingkan kandidat dengan data proyek sebelum memutuskan apakah transaksi baru akan digunakan.</p></div>
          </div>
          <div className="bcp-registry-toolbar">
            <select value={duplicateFilter} onChange={e => setDuplicateFilter(e.target.value)}>
              <option value="ALL">Semua kandidat</option>
              <option value="UNRESOLVED">Belum diputuskan</option>
              <option value="DECIDED">Sudah diputuskan</option>
              <option value="DOCUMENT">Dokumen</option>
              <option value="FILE">File sumber</option>
            </select>
            <span>{duplicateCandidates.length} kandidat ditemukan</span>
          </div>
          {duplicateCandidates.filter(row => {
            if (duplicateFilter === 'UNRESOLVED') return !row.decision;
            if (duplicateFilter === 'DECIDED') return !!row.decision;
            if (duplicateFilter === 'DOCUMENT' || duplicateFilter === 'FILE') return row.candidate_type === duplicateFilter;
            return true;
          }).length === 0 ? (
            <div className="bcp-empty"><div>✓</div><strong>Tidak ada kandidat duplikasi</strong><span>Kandidat akan muncul setelah file CEISA diperiksa pada halaman Upload.</span></div>
          ) : duplicateCandidates.filter(row => {
            if (duplicateFilter === 'UNRESOLVED') return !row.decision;
            if (duplicateFilter === 'DECIDED') return !!row.decision;
            if (duplicateFilter === 'DOCUMENT' || duplicateFilter === 'FILE') return row.candidate_type === duplicateFilter;
            return true;
          }).map(candidate => (
            <article className="bcp-stage-file" key={candidate.candidate_id}>
              <div className="bcp-stage-file-head">
                <div><strong>{candidate.candidate_type === 'FILE' ? candidate.file_name : `${candidate.document_type} · ${candidate.aju_number || 'Tanpa Nomor Aju'}`}</strong><small>{candidate.match_type} · {candidate.confidence} · {candidate.file_name}</small></div>
                <span className={'bcp-stage-badge ' + (candidate.decision ? 'ready' : 'warning')}>{candidate.decision?.decision || 'NEEDS DECISION'}</span>
              </div>
              <div className="bcp-stage-reason">{candidate.reason}</div>
              {candidate.candidate_type === 'DOCUMENT' ? (
                <>
                  <div className="bcp-info-grid">
                    <div><span>Data staging</span><strong>{candidate.aju_number || '—'} · {candidate.registration_number || '—'}</strong></div>
                    <div><span>Dokumen existing</span><strong>{candidate.existing_aju_number || candidate.existing_document_id || 'Dokumen lain dalam batch'}</strong></div>
                  </div>
                  <div className="bcp-import-actions">
                    <span>Setiap keputusan membutuhkan alasan dan tercatat di Audit Trail.</span>
                    <div className="bcp-heading-actions">
                      <button className="bcp-button bcp-button-secondary" onClick={() => decideDuplicate(candidate, 'USE_EXISTING')}>Use Existing</button>
                      <button className="bcp-button bcp-button-secondary" onClick={() => decideDuplicate(candidate, 'USE_NEW')}>Use New</button>
                      <button className="bcp-button bcp-button-secondary" onClick={() => decideDuplicate(candidate, 'KEEP_BOTH')}>Keep Both</button>
                      <button className="bcp-button bcp-button-quiet" onClick={() => decideDuplicate(candidate, 'DEFER_REVIEW')}>Defer Review</button>
                    </div>
                  </div>
                </>
              ) : (
                <div className="bcp-security-note"><strong>File duplikat persis</strong><span>SHA-256 file ini sudah terdaftar. File tidak akan diimpor ulang sebagai sumber baru. Periksa apakah dokumen terkait sudah tercatat dalam proyek.</span></div>
              )}
              {candidate.decision && <div className="bcp-stage-issues"><b>Alasan keputusan:</b> {candidate.decision.reason}<br/><b>Waktu:</b> {formatDate(candidate.decision.decided_at)}</div>}
            </article>
          ))}
          <div className="bcp-security-note"><strong>Aturan Duplicate Review</strong><span>Use Existing menautkan file sumber ke record lama tanpa membuat transaksi baru. Use New dan Keep Both mempertahankan record lama serta mencatat hubungan antar dokumen. Defer Review menahan dokumen dari Partial Import. Keputusan tidak menghapus data sumber.</span></div>
        </section>
      )}

      {view === 'documents' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">DOCUMENT REGISTRY</span><h2>Daftar Dokumen</h2><p>Satu record untuk setiap dokumen CEISA. Identitas dokumen dan detail sumber akan ditampilkan setelah proses impor tersedia.</p></div></div>
          <div className="bcp-registry-toolbar"><input value={documentQuery} onChange={e => setDocumentQuery(e.target.value)} placeholder="Cari Nomor Aju, Nomor Daftar, jenis BC, atau pihak..."/><span>{documentRows.length.toLocaleString('id-ID')} dokumen</span></div>
          {documentRows.length === 0 ? <div className="bcp-empty"><div>▤</div><strong>Belum ada dokumen</strong><span>Unggah dan konfirmasi dokumen CEISA melalui Partial Import untuk menampilkan record.</span></div> : (
            <div className="bcp-registry-table-wrap"><table className="bcp-registry-table"><thead><tr><th>Jenis BC</th><th>Nomor Daftar</th><th>Nomor Aju</th><th>Status</th><th>Amendment</th><th>Detail</th></tr></thead><tbody>
              {documentRows.map(row => <tr key={row.document_id}><td>{row.document_type}</td><td>{row.registration_number || row.current_data?.registration_number || '—'}</td><td>{row.aju_number || row.current_data?.nomor_aju || row.current_data?.aju_number || '—'}</td><td><span className="bcp-table-status">{row.business_status || '—'}</span></td><td>{row.amendment_status || 'ORIGINAL'}</td><td><button className="bcp-text-button" onClick={() => setSelectedDocumentId(row.document_id)}>Lihat</button></td></tr>)}
            </tbody></table></div>
          )}
          {selectedDocument && <div className="bcp-detail-panel"><div className="bcp-section-title"><h3>Detail Dokumen</h3><button className="bcp-text-button" onClick={() => setSelectedDocumentId(null)}>Tutup detail</button></div><div className="bcp-info-grid"><div><span>Document ID</span><code>{selectedDocument.document_id}</code></div><div><span>Nomor Daftar</span><strong>{selectedDocument.registration_number || '—'}</strong></div><div><span>Nomor Aju</span><strong>{selectedDocument.aju_number || selectedDocument.current_data?.nomor_aju || selectedDocument.current_data?.aju_number || '—'}</strong></div><div><span>Jumlah Item</span><strong>{selectedDocument.items.length}</strong></div></div><label className="bcp-period-assignment">Periode Bulanan<select value={selectedDocument.period_id || ''} onChange={e => e.target.value && changeDocumentPeriod(selectedDocument.document_id, e.target.value)}><option value="">Belum ditetapkan</option>{project.data.periods.filter(period => period.status !== 'CLOSED' || period.period_id === selectedDocument.period_id).map(period => <option key={period.period_id} value={period.period_id}>{period.period_key} · {period.status}</option>)}</select></label></div>}
        </section>
      )}

      {view === 'items' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">ITEM REGISTRY</span><h2>Daftar Barang</h2><p>Satu record untuk setiap seri barang. Riwayat transaksi item tidak diperlakukan sebagai saldo stok.</p></div></div>
          <div className="bcp-registry-toolbar"><input value={itemQuery} onChange={e => setItemQuery(e.target.value)} placeholder="Cari kode barang, uraian, atau spesifikasi..."/><span>{itemRows.length.toLocaleString('id-ID')} item</span></div>
          {itemRows.length === 0 ? <div className="bcp-empty"><div>▧</div><strong>Belum ada detail barang</strong><span>Detail barang akan muncul setelah dokumen CEISA berhasil diimpor.</span></div> : (
            <div className="bcp-registry-table-wrap"><table className="bcp-registry-table"><thead><tr><th>Seri</th><th>Kode Barang</th><th>Uraian</th><th>Jumlah</th><th>Satuan</th><th>Document ID</th></tr></thead><tbody>
              {itemRows.map(row => <tr key={row.item_id}><td>{row.item_sequence ?? '—'}</td><td>{row.item_code || '—'}</td><td>{row.description || '—'}</td><td>{row.quantity ?? '—'}</td><td>{row.unit || '—'}</td><td><button className="bcp-text-button" onClick={() => { setSelectedDocumentId(row.document_id); setView('documents'); }}>{row.document_id.slice(0, 12)}…</button></td></tr>)}
            </tbody></table></div>
          )}
        </section>
      )}

      {view === 'settings' && project && (
        <section className="bcp-content bcp-settings">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">LOCAL PROJECT PROFILE</span><h2>Pengaturan Proyek</h2><p>Identitas operator lokal hanya label tampilan, bukan autentikasi identitas pengguna.</p></div></div>
          <label>Nama Operator Lokal<input value={project.settings.operator.display_name} onChange={e => updateOperator(e.target.value)} maxLength={80}/></label>
          <label>Nama Proyek<input value={project.manifest.project_name} onChange={e => setProject(current => refreshProjectCounts({ ...current, manifest: { ...current.manifest, project_name: e.target.value } }))} maxLength={100}/></label>
          <label>Nama Perusahaan<input value={project.manifest.company_name || ''} onChange={e => setProject(current => refreshProjectCounts({ ...current, manifest: { ...current.manifest, company_name: e.target.value } }))} maxLength={150}/></label>
          <div className="bcp-security-note"><strong>Perhatian</strong><span>Perubahan pengaturan tersimpan ke Auto-Recovery. Klik Simpan .bcp untuk memperbarui file permanen.</span></div>
        </section>
      )}

      {view === 'diagnostics' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">LOCAL ENVIRONMENT</span><h2>System Diagnostics</h2><p>Pemeriksaan dasar kemampuan browser dan integritas struktur proyek.</p></div></div>
          <div className="bcp-diagnostic-list">
            <div><span>IndexedDB</span><strong className={'bcp-diagnostic-' + ('indexedDB' in window)}>{'indexedDB' in window ? 'Tersedia' : 'Tidak tersedia'}</strong></div>
            <div><span>Web Crypto API</span><strong className={'bcp-diagnostic-' + (!!window.crypto?.subtle)}>{window.crypto?.subtle ? 'Tersedia' : 'Tidak tersedia'}</strong></div>
            <div><span>Project Schema</span><strong className={validationErrors.length ? 'bcp-diagnostic-false' : 'bcp-diagnostic-true'}>{validationErrors.length ? 'Tidak valid' : 'Valid'}</strong></div>
            <div><span>Project Format</span><strong>{project.manifest.format} · {project.manifest.format_version}</strong></div>
            <div><span>Auto-Recovery</span><strong>{recoveryState === 'saved' ? 'Snapshot tersedia' : recoveryState === 'error' ? 'Gagal' : 'Menyiapkan'}</strong></div>
          </div>
          {validationErrors.length > 0 && <div className="bcp-error">{validationErrors.map(item => <div key={item}>{item}</div>)}</div>}
          <p className="bcp-muted">Integrity check memeriksa struktur proyek, checksum paket, dan file sumber yang disertakan.</p>
        </section>
      )}

      {showCreate && (
        <div className="bcp-modal-backdrop" role="presentation" onMouseDown={e => { if (e.target === e.currentTarget) setShowCreate(false); }}>
          <section className="bcp-modal" role="dialog" aria-modal="true" aria-labelledby="bcp-create-title">
            <span className="bcp-eyebrow">NEW PROJECT</span><h2 id="bcp-create-title">Buat Proyek BCP</h2>
            <label>Nama Proyek<input autoFocus value={projectName} onChange={e => setProjectName(e.target.value)} maxLength={100} placeholder="Contoh: BCP 2026"/></label>
            <label>Nama Perusahaan<input value={companyName} onChange={e => setCompanyName(e.target.value)} maxLength={150}/></label>
            <div className="bcp-modal-note">Proyek akan dibuat di browser. Simpan sebagai file .bcp untuk penyimpanan permanen.</div>
            <div className="bcp-modal-actions"><button className="bcp-button bcp-button-secondary" onClick={() => setShowCreate(false)}>Batal</button><button className="bcp-button bcp-button-primary" onClick={createProject}>Buat Proyek</button></div>
          </section>
        </div>
      )}
      <footer className="bcp-footer">REZA SYSTEM · BCP MODULE <span>LOCAL PROCESSING · NO SERVER STORAGE</span></footer>
    </main>
  );
}
