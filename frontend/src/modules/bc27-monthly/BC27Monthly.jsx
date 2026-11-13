import React, { useMemo, useState } from 'react';
import { DATABASE_HEADERS, OUTPUT_HEADERS, CATEGORY_MASTER, parseMonthlyBc27Files, exportMonthlyWorkbook } from './bc27MonthlyEngine';

const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const fmt = value => Number(value || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 });

export default function BC27Monthly() {
  const now = new Date();
  const [files, setFiles] = useState([]);
  const [period, setPeriod] = useState({ month: String(now.getMonth()+1).padStart(2,'0'), year: String(now.getFullYear()) });
  const [result, setResult] = useState(null);
  const [tab, setTab] = useState('database');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');

  const years = useMemo(() => {
    const current = now.getFullYear();
    return Array.from({length: 7}, (_,i) => String(current - 3 + i));
  }, []);
  const addFiles = list => {
    const incoming = [...list].filter(file => /\.xlsx$/i.test(file.name));
    setFiles(previous => {
      const map = new Map(previous.map(file => [`${file.name}:${file.size}:${file.lastModified}`, file]));
      incoming.forEach(file => map.set(`${file.name}:${file.size}:${file.lastModified}`, file));
      return [...map.values()];
    });
    setResult(null); setError(''); setNotice('');
    if (!incoming.length) setError('Pilih file Excel CEISA dengan format .xlsx.');
  };
  const processFiles = async () => {
    setBusy(true); setError(''); setNotice('');
    try {
      const parsed = await parseMonthlyBc27Files(files, period);
      setResult(parsed); setTab('database');
      const s = parsed.summary;
      setNotice(`Selesai membaca ${s.files} file. ${s.databaseRows.toLocaleString('id-ID')} baris barang masuk DATABASE dan ${s.outputRows.toLocaleString('id-ID')} kelompok barang terbentuk di OUTPUT.`);
      if (parsed.errors.length) setError(parsed.errors.join('\n'));
    } catch (e) {
      setResult(null); setError(e.message || 'Gagal memproses Excel CEISA.');
    } finally { setBusy(false); }
  };
  const reset = () => { setFiles([]); setResult(null); setError(''); setNotice(''); setTab('database'); };
  const exportExcel = () => {
    if (!result) return;
    exportMonthlyWorkbook(result.database, result.output, period);
  };
  const previewRows = tab === 'database' ? result?.database || [] : result?.output || [];
  const headers = tab === 'database' ? DATABASE_HEADERS : OUTPUT_HEADERS;
  const shownRows = previewRows.slice(0, 500);

  return <main className="bc27-page">
    <div className="bc27-heading">
      <div><div className="eyebrow">REZA SYSTEM · CEISA DATA PROCESSING</div><h1>BC 2.7 Monthly Report</h1><p>Rekapitulasi barang keluar Kawasan Berikat berdasarkan Kode Barang, Jenis Satuan, dan Kategori Barang.</p></div>
      <span className="bc27-local">PROSES LOKAL · BROWSER</span>
    </div>

    <section className="bc27-card">
      <div className="bc27-section-head"><div><h2>1. Periode laporan</h2><p>Transaksi dipilih berdasarkan Tanggal Daftar pada sheet HEADER kolom CQ.</p></div><span className="bc27-step">01</span></div>
      <div className="bc27-period">
        <label>Bulan<select value={period.month} onChange={e=>{setPeriod(p=>({...p,month:e.target.value}));setResult(null);setNotice('')}}>{monthNames.map((name,i)=><option key={name} value={String(i+1).padStart(2,'0')}>{name}</option>)}</select></label>
        <label>Tahun<select value={period.year} onChange={e=>{setPeriod(p=>({...p,year:e.target.value}));setResult(null);setNotice('')}}>{years.map(y=><option key={y} value={y}>{y}</option>)}</select></label>
      </div>
    </section>

    <section className="bc27-card">
      <div className="bc27-section-head"><div><h2>2. Upload Excel CEISA</h2><p>Pilih beberapa file sekaligus. Sistem membaca sheet HEADER, ENTITAS, dan BARANG.</p></div><span className="bc27-step">02</span></div>
      <label className="bc27-drop"><input type="file" accept=".xlsx" multiple onChange={e=>addFiles(e.target.files)}/><span className="bc27-upload-icon">⇧</span><b>Pilih beberapa file Excel CEISA (.xlsx)</b><small>BC 2.7 · Multiple Upload</small><span className="bc27-drop-note">File tidak dikirim ke server dan diproses langsung di browser.</span></label>
      {files.length>0&&<div className="bc27-file-list">{files.map(file=><div key={`${file.name}:${file.size}:${file.lastModified}`}><span className="bc27-file-icon">XLSX</span><b>{file.name}</b><small>{(file.size/1024).toLocaleString('id-ID',{maximumFractionDigits:0})} KB</small><button onClick={()=>{setFiles(prev=>prev.filter(x=>x!==file));setResult(null);setNotice('')}} aria-label={`Hapus ${file.name}`}>×</button></div>)}</div>}
      <div className="bc27-actions"><button className="bc27-primary" disabled={!files.length||busy} onClick={processFiles}>{busy?'MEMBACA DAN MENGHITUNG...':'PROSES DATA BC 2.7'}</button><button className="bc27-secondary" disabled={!files.length&&!result} onClick={reset}>Reset</button></div>
      {error&&<div className="bc27-error">{error}</div>}{notice&&<div className="bc27-notice">{notice}</div>}
    </section>

    {result&&<>
      <section className="bc27-card">
        <div className="bc27-section-head"><div><h2>3. Ringkasan hasil pembacaan</h2><p>Dokumen BC 2.7 keluar diidentifikasi dari kode entitas 3 = Chinli.</p></div><span className="bc27-step">03</span></div>
        <div className="bc27-stats">
          {[
            ['File dibaca',result.summary.files],['Dokumen BC 2.7',result.summary.totalHeaderDocs],['BC 2.7 Keluar',result.summary.outgoingDocs],['Baris DATABASE',result.summary.databaseRows],['Kelompok OUTPUT',result.summary.outputRows],['Duplikat AJU dilewati',result.summary.duplicateDocs],['Bukan BC 2.7 Keluar',result.summary.nonOutgoingDocs],['Di luar periode',result.summary.outsidePeriodDocs]
          ].map(([label,val])=><div key={label}><small>{label}</small><b>{Number(val).toLocaleString('id-ID')}</b></div>)}
        </div>
        {result.errors.length>0&&<div className="bc27-warning"><b>File yang perlu diperiksa</b>{result.errors.map((x,i)=><p key={i}>{x}</p>)}</div>}
        {result.notices.length>0&&<details className="bc27-details"><summary>{result.notices.length} catatan validasi data</summary><ul>{result.notices.slice(0,100).map((x,i)=><li key={i}>{x}</li>)}</ul>{result.notices.length>100&&<small>Menampilkan 100 catatan pertama.</small>}</details>}
      </section>

      <section className="bc27-card">
        <div className="bc27-section-head"><div><h2>4. DATABASE dan OUTPUT</h2><p>DATABASE menyimpan rincian per barang per AJU. OUTPUT menjumlahkan transaksi berdasarkan tiga kunci penggabungan.</p></div><span className="bc27-step">04</span></div>
        <div className="bc27-tabs"><button className={tab==='database'?'active':''} onClick={()=>setTab('database')}>DATABASE ({result.database.length.toLocaleString('id-ID')})</button><button className={tab==='output'?'active':''} onClick={()=>setTab('output')}>OUTPUT ({result.output.length.toLocaleString('id-ID')})</button></div>
        {previewRows.length===0?<div className="bc27-empty">Tidak ada data untuk periode yang dipilih.</div>:<>
          <div className="bc27-table-wrap"><table className="bc27-table"><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{shownRows.map((row,i)=><tr key={`${tab}-${i}`}>{headers.map(h=><td key={h} className={h==='Uraian Barang'||h==='Entitas Perusahaan Tujuan'?'wide-cell':''}>{h==='Jumlah Satuan'?fmt(row[h]):row[h]??''}</td>)}</tr>)}</tbody></table></div>
          {previewRows.length>shownRows.length&&<p className="bc27-preview-note">Pratinjau menampilkan {shownRows.length.toLocaleString('id-ID')} dari {previewRows.length.toLocaleString('id-ID')} baris. Ekspor Excel tetap memuat seluruh data.</p>}
        </>}
        <div className="bc27-export-row"><button className="bc27-primary" disabled={!result.database.length} onClick={exportExcel}>⇩ EXPORT EXCEL · DATABASE + OUTPUT</button><small>File Excel berisi dua sheet: DATABASE dan OUTPUT.</small></div>
      </section>
    </>}

    <section className="bc27-card bc27-rules">
      <h2>Aturan pembacaan dan penggabungan</h2>
      <ul>
        <li>Hanya dokumen berkode 27 dengan entitas kode 3 bernama persis <b>CHINLI PLASTIC MATERIALS INDONESIA</b> yang diproses sebagai BC 2.7 Keluar.</li>
        <li>Entitas tujuan dibaca dari kode entitas 8.</li>
        <li>Nomor AJU, nomor daftar, dan tanggal daftar berasal dari HEADER kolom A, CP, dan CQ.</li>
        <li>Kode barang, HS Code, uraian, satuan, jumlah, dan kategori berasal dari BARANG kolom D, C, E, J, K, dan AW.</li>
        <li>Duplikat Nomor AJU antarfile hanya dihitung satu kali; file pertama yang diproses menjadi sumber data.</li>
        <li>OUTPUT menggabungkan berdasarkan Kode Barang + Jenis Satuan + Kategori Barang.</li>
      </ul>
      <div className="bc27-category"><b>Master kategori CEISA</b><div>{Object.entries(CATEGORY_MASTER).map(([code,name])=><span key={code}><b>{code}</b> {name}</span>)}</div></div>
    </section>
  </main>;
}
