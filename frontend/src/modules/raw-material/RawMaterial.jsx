import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { BAHANBAKU_HEADERS, BAHANBAKUTARIF_HEADERS, parseCeisaFile, buildBahanBakuRows, buildTarifRows, rowsToAoa } from './rawMaterialEngine';

const initialSettings = { bm: 15, ppn: 11, pph: 2.5, ppnLokal: 11, kodeTarif: '1', kodeFasilitas: '3', tarifFasilitas: 100 };
const fmt = n => Number(n || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 });
const copyRows = async (headers, rows) => {
  const text = [headers, ...rows.map(r => headers.map(h => r[h] ?? ''))].map(r => r.map(v => String(v ?? '').replace(/[\t\r\n]/g, ' ')).join('\t')).join('\n');
  await navigator.clipboard.writeText(text);
};

export default function RawMaterial() {
  const [files, setFiles] = useState([]), [docs, setDocs] = useState([]), [selected, setSelected] = useState({}), [edits, setEdits] = useState({});
  const [settings, setSettings] = useState(initialSettings), [busy, setBusy] = useState(false), [error, setError] = useState(''), [notice, setNotice] = useState('');
  const [tab, setTab] = useState('bahanbaku');
  const allItems = useMemo(() => docs.flatMap(d => d.items.map(item => ({ ...item, docAju: d.aju, docType: d.docType, docCode: d.docCode }))), [docs]);
  const getItem = item => ({ ...item, ...edits[item.id], quantity: edits[item.id]?.quantity ?? item.quantity, series: edits[item.id]?.series ?? item.series, seriesBahan: edits[item.id]?.seriesBahan ?? 1 });
  const chosen = allItems.filter(x => selected[x.id]).map(getItem);
  const selectedByDoc = docs.map(doc => ({ doc, items: doc.items.filter(x => selected[x.id]).map(getItem) })).filter(x => x.items.length);
  const bahanRows = selectedByDoc.flatMap(({doc,items}) => buildBahanBakuRows(doc, items));
  const tarifRows = selectedByDoc.flatMap(({doc,items}) => buildTarifRows(doc, items, settings));
  const updateEdit = (id, key, value) => setEdits(prev => ({...prev, [id]: {...prev[id], [key]: value}}));
  const processFiles = async () => {
    setBusy(true); setError(''); setNotice('');
    try {
      const parsed = await Promise.all(files.map(parseCeisaFile));
      const seen = new Set();
      parsed.forEach(d => { if (seen.has(d.aju)) throw new Error(`Nomor AJU ${d.aju} terdeteksi lebih dari satu kali.`); seen.add(d.aju); });
      setDocs(parsed); setSelected({}); setEdits({}); setNotice(`${parsed.length} dokumen berhasil dibaca. Total ${parsed.reduce((n,d)=>n+d.items.length,0)} barang ditemukan.`);
    } catch (e) { setError(e.message || 'Gagal membaca file CEISA.'); setDocs([]); setSelected({}); }
    finally { setBusy(false); }
  };
  const setAll = checked => { const next = {}; allItems.forEach(x => { next[x.id] = checked; }); setSelected(next); };
  const clear = () => { setFiles([]); setDocs([]); setSelected({}); setEdits({}); setError(''); setNotice(''); };
  const handleCopy = async (headers, rows, label) => { try { await copyRows(headers, rows); setNotice(`${label} berhasil disalin. Silakan paste ke Excel.`); } catch { setError('Tidak dapat menyalin otomatis. Pastikan browser mengizinkan akses clipboard.'); } };
  const exportExcel = (name, headers, rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rowsToAoa(headers, rows)); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, name); XLSX.writeFile(wb, `${name}.xlsx`);
  };

  return <main className="raw-page">
    <div className="raw-heading"><div><div className="eyebrow">REZA SYSTEM · CEISA DATA PROCESSING</div><h1>Pengambilan Data Bahan Baku</h1><p>Ambil data barang dari sheet HEADER dan BARANG, lalu bentuk output BAHANBAKU dan BAHANBAKUTARIF.</p></div><span className="raw-local">PROSES LOKAL</span></div>
    <section className="raw-card">
      <div className="raw-section-title"><div><h2>1. Upload Excel CEISA</h2><p>Hanya sheet HEADER dan BARANG yang dibaca. Sheet lainnya diabaikan.</p></div><span className="raw-step">01</span></div>
      <label className="raw-drop"><input type="file" accept=".xlsx" multiple onChange={e=>{setFiles([...e.target.files]);setError('');setNotice('')}}/><b>Pilih satu atau beberapa file CEISA (.xlsx)</b><span>BC 2.7 · BC 2.3 · BC 4.0</span></label>
      {files.length>0&&<div className="raw-file-list">{files.map(f=><div key={f.name}>{f.name}<small>{(f.size/1024).toFixed(1)} KB</small></div>)}</div>}
      <div className="raw-actions"><button className="raw-primary" disabled={!files.length||busy} onClick={processFiles}>{busy?'MEMBACA EXCEL...':'BACA DATA CEISA'}</button>{(docs.length||files.length>0)&&<button className="raw-secondary" onClick={clear}>Reset</button>}</div>
      {error&&<div className="raw-error">{error}</div>}{notice&&<div className="raw-notice">{notice}</div>}
    </section>

    {docs.length>0&&<>
      <section className="raw-card">
        <div className="raw-section-title"><div><h2>2. Pilih barang bahan baku</h2><p>Pilihan barang berlaku sama untuk kedua output. Kolom seri dan jumlah dapat diedit.</p></div><span className="raw-step">02</span></div>
        <div className="raw-meta">{docs.map(d=><span key={d.aju}>{d.docType} · AJU {d.aju} · Daftar {d.nomorDaftar || '—'} · NDPBM {d.ndpbm}</span>)}</div>
        <div className="raw-table-scroll raw-item-scroll"><table className="raw-table raw-items"><thead><tr><th className="sticky-check"><input type="checkbox" aria-label="Pilih semua barang" checked={allItems.length>0&&chosen.length===allItems.length} onChange={e=>setAll(e.target.checked)}/><br/>Pilih semua</th><th className="sticky-series">Seri Barang</th><th>Seri Bahan Baku</th><th>AJU</th><th>Jenis</th><th>HS</th><th>Kode Barang</th><th>Uraian</th><th>Kode Satuan</th><th>Jumlah Satuan</th><th>CIF</th><th>NDPBM</th></tr></thead><tbody>{allItems.map(item=>{const v=getItem(item);return <tr key={item.id}><td className="sticky-check"><input type="checkbox" checked={!!selected[item.id]} onChange={e=>setSelected(p=>({...p,[item.id]:e.target.checked}))}/></td><td className="sticky-series"><input className="raw-mini" value={v.series} onChange={e=>updateEdit(item.id,'series',e.target.value)}/></td><td><input className="raw-mini" value={v.seriesBahan} onChange={e=>updateEdit(item.id,'seriesBahan',e.target.value)}/></td><td>{item.docAju}</td><td>{item.docType}</td><td>{item.hs}</td><td>{item.kodeBarang}</td><td className="raw-description">{item.uraian}</td><td>{item.kodeSatuan}</td><td><input className="raw-number" type="number" min="0" value={v.quantity} onChange={e=>updateEdit(item.id,'quantity',e.target.value)}/></td><td>{fmt(item.sourceCif)}</td><td>{fmt(item.ndpbm)}</td></tr>})}</tbody></table></div>
        <div className="raw-selected">{chosen.length} dari {allItems.length} barang dipilih</div>
      </section>

      <section className="raw-card">
        <div className="raw-section-title"><div><h2>3. Pengaturan tarif</h2><p>Default dapat diubah sebelum output dibuat.</p></div><span className="raw-step">03</span></div>
        <div className="raw-settings">
          {[['bm','BM (%)'],['ppn','PPN (%)'],['pph','PPH (%)'],['ppnLokal','PPNLOKAL (%)'],['kodeTarif','Kode Tarif'],['kodeFasilitas','Kode Fasilitas'],['tarifFasilitas','Tarif Fasilitas']].map(([k,label])=><label key={k}>{label}<input type="number" step="any" value={settings[k]} onChange={e=>setSettings(s=>({...s,[k]:e.target.value}))}/></label>)}
        </div>
      </section>

      <section className="raw-card">
        <div className="raw-section-title"><div><h2>4. Preview output</h2><p>Data kedua output berasal dari barang terpilih yang sama.</p></div><span className="raw-step">04</span></div>
        <div className="raw-tabs"><button className={tab==='bahanbaku'?'active':''} onClick={()=>setTab('bahanbaku')}>BAHANBAKU ({bahanRows.length})</button><button className={tab==='tarif'?'active':''} onClick={()=>setTab('tarif')}>BAHANBAKUTARIF ({tarifRows.length})</button></div>
        {chosen.length===0?<div className="raw-empty">Pilih minimal satu barang untuk menampilkan preview.</div>:tab==='bahanbaku'?<>
          <div className="raw-preview-scroll"><table className="raw-table raw-preview"><thead><tr>{BAHANBAKU_HEADERS.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{bahanRows.map((r,i)=><tr key={i}>{BAHANBAKU_HEADERS.map(h=><td key={h}>{r[h]=== ''?'':typeof r[h]==='number'?fmt(r[h]):r[h]}</td>)}</tr>)}</tbody></table></div>
          <div className="raw-output-actions"><button className="raw-primary" onClick={()=>handleCopy(BAHANBAKU_HEADERS,bahanRows,'BAHANBAKU')}>SALIN DATA BAHANBAKU</button><button className="raw-secondary" onClick={()=>exportExcel('BAHANBAKU',BAHANBAKU_HEADERS,bahanRows)}>Unduh Excel</button></div>
        </>:<>
          <div className="raw-preview-scroll"><table className="raw-table raw-preview"><thead><tr>{BAHANBAKUTARIF_HEADERS.map((h,i)=><th key={`${h}-${i}`}>{h}</th>)}</tr></thead><tbody>{tarifRows.map((r,i)=><tr key={i}>{BAHANBAKUTARIF_HEADERS.map((h,j)=><td key={`${h}-${j}`}>{r[h]=== ''?'':r[h]}</td>)}</tr>)}</tbody></table></div>
          <div className="raw-output-actions"><button className="raw-primary" onClick={()=>handleCopy(BAHANBAKUTARIF_HEADERS,tarifRows,'BAHANBAKUTARIF')}>SALIN DATA BAHANBAKUTARIF</button><button className="raw-secondary" onClick={()=>exportExcel('BAHANBAKUTARIF',BAHANBAKUTARIF_HEADERS,tarifRows)}>Unduh Excel</button></div>
        </>}
        <p className="raw-footnote">NILAI BAYAR, NILAI FASILITAS, dan NILAI SUDAH DILUNASI selalu dikosongkan. NETTO dan BRUTO juga dikosongkan pada output BAHANBAKU.</p>
      </section>
    </>}
  </main>;
}
