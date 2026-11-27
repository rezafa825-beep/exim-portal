import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import './bcp.css';
import { createEmptyProject, refreshProjectCounts, validateProject } from './projectSchema';
import { deleteRecovery, getRecovery, listRecentProjects, saveRecentProject, saveRecovery } from './projectStorage';
import { downloadProjectArchive, readProjectArchive } from './projectPackage';
import { getDocumentView, searchDocuments, searchItems } from './registryEngine';

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
        {project && <button className={view === 'documents' ? 'active' : ''} onClick={() => { setSelectedDocumentId(null); go('documents'); }}>Document Registry</button>}
        {project && <button className={view === 'items' ? 'active' : ''} onClick={() => go('items')}>Item Registry</button>}
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
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">PROJECT OVERVIEW</span><h2>Dashboard BCP</h2><p>Ringkasan proyek saat ini. Modul pencatatan dan ekstraksi CEISA akan dibangun pada milestone berikutnya.</p></div></div>
          <div className="bcp-metric-grid">
            <div><span>Dokumen</span><strong>{project.data.documents.length.toLocaleString('id-ID')}</strong><small>Document Registry</small></div>
            <div><span>Detail Barang</span><strong>{project.data.items.length.toLocaleString('id-ID')}</strong><small>Item Registry</small></div>
            <div><span>File Sumber</span><strong>{project.sources.source_registry.length.toLocaleString('id-ID')}</strong><small>Source Registry</small></div>
            <div><span>Riwayat Aktivitas</span><strong>{project.data.audit_log.length.toLocaleString('id-ID')}</strong><small>Audit Trail lokal</small></div>
          </div>
          <div className="bcp-next-stage">
            <span className="bcp-eyebrow">DEVELOPMENT ROADMAP</span>
            <h3>Fondasi proyek berhasil disiapkan</h3>
            <p>Project Engine dan penyimpanan pemulihan lokal aktif. Tahap berikutnya adalah Registry, lalu Unified Upload dan parser BC 2.7 IN/OUT.</p>
            <div className="bcp-progress"><span style={{ width: '12.5%' }}/></div>
            <small>M1 · BCP Foundation — implementasi awal</small>
          </div>
          <div className="bcp-section-title"><h3>Informasi Proyek</h3></div>
          <div className="bcp-info-grid"><div><span>Project ID</span><code>{project.manifest.project_id}</code></div><div><span>Format</span><strong>{project.manifest.format} · {project.manifest.format_version}</strong></div><div><span>Schema Version</span><strong>{project.manifest.schema_version}</strong></div><div><span>Pembaruan</span><strong>{formatDate(project.manifest.updated_at)}</strong></div></div>
        </section>
      )}

      {view === 'documents' && project && (
        <section className="bcp-content">
          <div className="bcp-section-heading"><div><span className="bcp-eyebrow">DOCUMENT REGISTRY</span><h2>Daftar Dokumen</h2><p>Satu record untuk setiap dokumen CEISA. Identitas dokumen dan detail sumber akan ditampilkan setelah proses impor tersedia.</p></div></div>
          <div className="bcp-registry-toolbar"><input value={documentQuery} onChange={e => setDocumentQuery(e.target.value)} placeholder="Cari Nomor Aju, Nomor Daftar, jenis BC, atau pihak..."/><span>{documentRows.length.toLocaleString('id-ID')} dokumen</span></div>
          {documentRows.length === 0 ? <div className="bcp-empty"><div>▤</div><strong>Belum ada dokumen</strong><span>Dokumen akan muncul di sini setelah CEISA Extraction dan Import Engine tersedia.</span></div> : (
            <div className="bcp-registry-table-wrap"><table className="bcp-registry-table"><thead><tr><th>Jenis BC</th><th>Nomor Daftar</th><th>Nomor Aju</th><th>Status</th><th>Amendment</th><th>Detail</th></tr></thead><tbody>
              {documentRows.map(row => <tr key={row.document_id}><td>{row.document_type}</td><td>{row.registration_number || row.current_data?.registration_number || '—'}</td><td>{row.aju_number || row.current_data?.nomor_aju || row.current_data?.aju_number || '—'}</td><td><span className="bcp-table-status">{row.business_status || '—'}</span></td><td>{row.amendment_status || 'ORIGINAL'}</td><td><button className="bcp-text-button" onClick={() => setSelectedDocumentId(row.document_id)}>Lihat</button></td></tr>)}
            </tbody></table></div>
          )}
          {selectedDocument && <div className="bcp-detail-panel"><div className="bcp-section-title"><h3>Detail Dokumen</h3><button className="bcp-text-button" onClick={() => setSelectedDocumentId(null)}>Tutup detail</button></div><div className="bcp-info-grid"><div><span>Document ID</span><code>{selectedDocument.document_id}</code></div><div><span>Nomor Daftar</span><strong>{selectedDocument.registration_number || '—'}</strong></div><div><span>Nomor Aju</span><strong>{selectedDocument.aju_number || selectedDocument.current_data?.nomor_aju || selectedDocument.current_data?.aju_number || '—'}</strong></div><div><span>Jumlah Item</span><strong>{selectedDocument.items.length}</strong></div></div></div>}
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
          <p className="bcp-muted">Integrity check pada tahap ini memeriksa struktur JSON dan format proyek. Verifikasi checksum sumber file akan ditambahkan saat modul Source Registry dan import dibangun.</p>
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
