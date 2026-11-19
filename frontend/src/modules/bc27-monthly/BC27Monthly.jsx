import React, { useEffect, useMemo, useState } from 'react';
import { DATABASE_HEADERS, OUTPUT_HEADERS, CATEGORY_MASTER, parseMonthlyBc27Files, exportMonthlyWorkbook } from './bc27MonthlyEngine';
import { parseErpMappings, parseFinishedGoodsInventory, buildCrosscheck, exportCrosscheck, downloadMappingTemplate } from './bc27CrosscheckEngine';
import { buildAdoniaSkuOutputs } from './bc27AdoniaSkuEngine';

const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const fmt = value => Number(value || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 });
const MASTER_STORAGE_KEY = 'rezasystem_bc27_master_mapping_v1';

export default function BC27Monthly() {
  const now = new Date();
  const [files, setFiles] = useState([]);
  const [period, setPeriod] = useState({ month: String(now.getMonth()+1).padStart(2,'0'), year: String(now.getFullYear()) });
  const [result, setResult] = useState(null);
  const [tab, setTab] = useState('database');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [erpFile, setErpFile] = useState(null);
  const [inventoryFile, setInventoryFile] = useState(null);
  const [mappings, setMappings] = useState(() => { try { const saved = window.localStorage.getItem(MASTER_STORAGE_KEY); return saved ? JSON.parse(saved) : []; } catch { return []; } });
  useEffect(() => { try { if (mappings.length) window.localStorage.setItem(MASTER_STORAGE_KEY, JSON.stringify(mappings)); else window.localStorage.removeItem(MASTER_STORAGE_KEY); } catch { /* Storage may be unavailable or full; current session still works. */ } }, [mappings]);
  const [inventory, setInventory] = useState(null);
  const [crosscheck, setCrosscheck] = useState(null);
  const [crossBusy, setCrossBusy] = useState(false);
  const [crossError, setCrossError] = useState('');
  const [crossFilter, setCrossFilter] = useState('ALL');
  const [mappingMode, setMappingMode] = useState('replace');
  const [manualCustomer, setManualCustomer] = useState('');
  const [manualCeisa, setManualCeisa] = useState('');
  const [manualInternal, setManualInternal] = useState('');
  const [invoiceFiles, setInvoiceFiles] = useState([]);
  const [skuResult, setSkuResult] = useState(null);
  const [skuBusy, setSkuBusy] = useState(false);
  const [skuError, setSkuError] = useState('');
  const [workflowOption, setWorkflowOption] = useState('choose');

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
    setResult(null); setError(''); setNotice(''); setInventoryFile(null); setInventory(null); setCrosscheck(null); setCrossError(''); setSkuResult(null); setSkuError('');
    if (!incoming.length) setError('Pilih file Excel CEISA dengan format .xlsx.');
  };
  const processFiles = async () => {
    setBusy(true); setError(''); setNotice('');
    try {
      const parsed = await parseMonthlyBc27Files(files, period);
      setResult(parsed); setTab('database'); setWorkflowOption('choose'); setInventoryFile(null); setInventory(null); setCrosscheck(null); setCrossError(''); setSkuResult(null); setSkuError('');
      const s = parsed.summary;
      setNotice(`Selesai membaca ${s.files} file. ${s.databaseRows.toLocaleString('id-ID')} baris barang masuk DATABASE dan ${s.outputRows.toLocaleString('id-ID')} kelompok barang terbentuk di OUTPUT.`);
      if (parsed.errors.length) setError(parsed.errors.join('\n'));
    } catch (e) {
      setResult(null); setError(e.message || 'Gagal memproses Excel CEISA.');
    } finally { setBusy(false); }
  };
  const reset = () => { setFiles([]); setResult(null); setWorkflowOption('choose'); setError(''); setNotice(''); setTab('database'); setErpFile(null); setInventoryFile(null); setInventory(null); setCrosscheck(null); setCrossError(''); setInvoiceFiles([]); setSkuResult(null); setSkuError(''); };
  const processAdoniaSku = async () => {
    setSkuBusy(true); setSkuError(''); setSkuResult(null);
    try {
      const parsed = await buildAdoniaSkuOutputs(files, invoiceFiles, period);
      setSkuResult(parsed);
    } catch (e) {
      setSkuError(e.message || 'Gagal mencocokkan SKU Invoice dengan CEISA.');
    } finally { setSkuBusy(false); }
  };
  const downloadAdoniaSkuFile = output => {
    const blob = new Blob([output.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url; anchor.download = output.name; anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const exportExcel = () => {
    if (!result) return;
    exportMonthlyWorkbook(result.database, result.output, period);
  };
  const previewRows = tab === 'database' ? result?.database || [] : result?.output || [];
  const headers = tab === 'database' ? DATABASE_HEADERS : OUTPUT_HEADERS;
  const shownRows = previewRows.slice(0, 500);

  return <main className="bc27-page">
    <div className="bc27-heading">
      <div><div className="eyebrow">REZA SYSTEM · CEISA DATA PROCESSING</div><h1>BC 2.7 Monthly Report</h1><p>Rekapitulasi barang keluar Kawasan Berikat. Adonia menggunakan Spesifikasi Lain sebagai identitas barang; perusahaan lain menggunakan Kode Barang.</p></div>
      <span className="bc27-local">PROSES LOKAL · BROWSER</span>
    </div>

    <section className="bc27-card">
      <div className="bc27-section-head"><div><h2>1. Periode laporan</h2><p>Transaksi dipilih berdasarkan Tanggal Daftar pada sheet HEADER kolom CQ.</p></div><span className="bc27-step">01</span></div>
      <div className="bc27-period">
        <label>Bulan<select value={period.month} onChange={e=>{setPeriod(p=>({...p,month:e.target.value}));setResult(null);setInventory(null);setCrosscheck(null);setCrossError('');setNotice('')}}>{monthNames.map((name,i)=><option key={name} value={String(i+1).padStart(2,'0')}>{name}</option>)}</select></label>
        <label>Tahun<select value={period.year} onChange={e=>{setPeriod(p=>({...p,year:e.target.value}));setResult(null);setInventory(null);setCrosscheck(null);setCrossError('');setNotice('')}}>{years.map(y=><option key={y} value={y}>{y}</option>)}</select></label>
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


      <section className="bc27-card bc27-workflow-choice">
        <div className="bc27-section-head"><div><h2>4. Pilih proses lanjutan</h2><p>Data CEISA sudah berhasil dibaca. Pilih salah satu alur yang ingin dikerjakan.</p></div><span className="bc27-step">04</span></div>
        <div className="bc27-workflow-options">
          <button type="button" className={workflowOption==='standard'?'selected':''} onClick={()=>setWorkflowOption('standard')}><span className="bc27-option-number">A</span><b>Export Monthly Report & Crosscheck</b><small>Lanjutkan hasil DATABASE + OUTPUT, export Excel, kemudian crosscheck dengan IT Inventory.</small><span className="bc27-option-action">PILIH ALUR INI →</span></button>
          <button type="button" className={workflowOption==='adonia'?'selected':''} onClick={()=>setWorkflowOption('adonia')}><span className="bc27-option-number">B</span><b>Add SKU to Spesifikasi Lain · Adonia</b><small>Upload satu atau beberapa Excel Invoice Adonia, cocokkan per Nomor Invoice dan Nomor Urut Item, lalu export file CEISA dengan SKU.</small><span className="bc27-option-action">PILIH ALUR INI →</span></button>
        </div>
        {workflowOption!=='choose'&&<div className="bc27-workflow-current">Alur dipilih: <b>{workflowOption==='standard'?'Export Monthly Report & Crosscheck IT vs CEISA':'Add SKU to Spesifikasi Lain · Adonia'}</b><button type="button" className="bc27-secondary" onClick={()=>setWorkflowOption('choose')}>GANTI OPSI</button></div>}
      </section>

      {workflowOption==='adonia'&&<section className="bc27-card bc27-adonia-sku">
        <div className="bc27-section-head"><div><h2>4. Add SKU to Spesifikasi Lain · Khusus Adonia</h2><p>Hasil akhir hanya memuat transaksi Adonia dalam satu sheet BARANG dengan delapan kolom ringkas. SKU Invoice digabungkan ke Spesifikasi Lain setelah pencocokan berhasil.</p></div><span className="bc27-step">04</span></div>
        <div className="bc27-sku-info"><b>Alur pencocokan otomatis</b><ol><li>Nomor AJU dan entitas tujuan Adonia diidentifikasi dari CEISA.</li><li>Nomor Invoice dibaca dari sheet DOKUMEN (kode dokumen 380, dengan 217 sebagai fallback).</li><li>Nomor urut Invoice dicocokkan dengan SERI BARANG CEISA.</li><li>MATERIAL CODE Invoice harus sama dengan KODE BARANG CEISA sebelum SKU dimasukkan ke SPESIFIKASI LAIN.</li></ol></div>
        <label className="bc27-drop bc27-invoice-drop"><input type="file" accept=".xlsx" multiple onChange={e=>{const incoming=[...e.target.files].filter(file=>/\.xlsx$/i.test(file.name));setInvoiceFiles(prev=>[...new Map([...prev,...incoming].map(file=>[file.name+':'+file.size+':'+file.lastModified,file])).values()]);setSkuResult(null);setSkuError('');e.target.value='';}}/><span className="bc27-upload-icon">⇧</span><b>Pilih beberapa Excel Invoice Adonia (.xlsx)</b><small>Invoice biasa maupun file gabungan beberapa sheet Invoice</small><span className="bc27-drop-note">Nomor Invoice digunakan untuk memisahkan transaksi; satu file boleh berisi beberapa sheet Invoice.</span></label>
        {invoiceFiles.length>0&&<div className="bc27-file-list">{invoiceFiles.map(file=><div key={file.name+':'+file.size+':'+file.lastModified}><span className="bc27-file-icon">XLSX</span><b>{file.name}</b><small>{(file.size/1024).toLocaleString('id-ID',{maximumFractionDigits:0})} KB</small><button onClick={()=>{setInvoiceFiles(prev=>prev.filter(x=>x!==file));setSkuResult(null)}} aria-label={'Hapus '+file.name}>×</button></div>)}</div>}
        <div className="bc27-actions"><button className="bc27-primary" disabled={!files.length||!invoiceFiles.length||skuBusy} onClick={processAdoniaSku}>{skuBusy?'MEMBACA INVOICE & MENCOCOKKAN SKU...':'PROSES ADD SKU ADONIA'}</button><button className="bc27-secondary" disabled={!invoiceFiles.length&&!skuResult} onClick={()=>{setInvoiceFiles([]);setSkuResult(null);setSkuError('')}}>Reset Invoice</button></div>
        {skuError&&<div className="bc27-error">{skuError}</div>}
        {skuResult&&<>
          <div className="bc27-stats bc27-sku-stats">
            {[
              ['File Invoice',skuResult.summary.invoiceFiles],['Sheet Invoice terbaca',skuResult.summary.invoiceSheetsRead],['AJU Adonia',skuResult.summary.adoniaAjuCount],['Item cocok invoice',skuResult.summary.itemMatched],['Item tidak cocok',skuResult.summary.itemUnmatched],['SKU sudah tercantum',skuResult.summary.itemSkipped],['Baris output',skuResult.summary.outputRows],['File output',skuResult.summary.outputFiles]
            ].map(([label,val])=><div key={label}><small>{label}</small><b>{Number(val).toLocaleString('id-ID')}</b></div>)}
          </div>
          {skuResult.issues.length>0&&<div className="bc27-warning"><b>Catatan pembacaan</b>{skuResult.issues.map((x,i)=><p key={i}>{x}</p>)}</div>}
          <div className="bc27-table-wrap"><table className="bc27-table"><thead><tr>{['Nomor AJU','Nomor Daftar','Tanggal Daftar','Kode Barang','Uraian Barang','Spesifikasi Lain','Jenis Satuan','Kuantiti'].map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{skuResult.outputs.flatMap(output=>output.rowsPreview||[]).slice(0,500).map((row,i)=><tr key={row['Nomor AJU']+'-'+i}>{['Nomor AJU','Nomor Daftar','Tanggal Daftar','Kode Barang','Uraian Barang','Spesifikasi Lain','Jenis Satuan','Kuantiti'].map(h=><td key={h} className={h==='Uraian Barang'||h==='Spesifikasi Lain'?'wide-cell':''}>{row[h]??''}</td>)}</tr>)}</tbody></table></div>
          <p className="bc27-preview-note">Pratinjau data tersedia setelah proses selesai. File Excel hasil berisi seluruh baris Adonia.</p>
          <div className="bc27-sku-downloads"><b>Download hasil Adonia · satu sheet BARANG · delapan kolom</b>{skuResult.outputs.map((output,i)=><div key={output.name}><span>{output.name} · {output.changed} SKU diisi</span><button className="bc27-secondary" onClick={()=>downloadAdoniaSkuFile(output)}>DOWNLOAD EXCEL</button></div>)}</div>
        </>}
      </section>}

      {workflowOption==='standard'&&<section className="bc27-card">
        <div className="bc27-section-head"><div><h2>5. DATABASE dan OUTPUT</h2><p>DATABASE menyimpan rincian per barang per AJU. OUTPUT menjumlahkan transaksi berdasarkan identitas barang, Jenis Satuan, dan Kategori Barang.</p></div><span className="bc27-step">05</span></div>
        <div className="bc27-tabs"><button className={tab==='database'?'active':''} onClick={()=>setTab('database')}>DATABASE ({result.database.length.toLocaleString('id-ID')})</button><button className={tab==='output'?'active':''} onClick={()=>setTab('output')}>OUTPUT ({result.output.length.toLocaleString('id-ID')})</button></div>
        {previewRows.length===0?<div className="bc27-empty">Tidak ada data untuk periode yang dipilih.</div>:<>
          <div className="bc27-table-wrap"><table className="bc27-table"><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{shownRows.map((row,i)=><tr key={`${tab}-${i}`}>{headers.map(h=><td key={h} className={h==='Uraian Barang'||h==='Entitas Perusahaan Tujuan'?'wide-cell':''}>{h==='Jumlah Satuan'?fmt(row[h]):row[h]??''}</td>)}</tr>)}</tbody></table></div>
          {previewRows.length>shownRows.length&&<p className="bc27-preview-note">Pratinjau menampilkan {shownRows.length.toLocaleString('id-ID')} dari {previewRows.length.toLocaleString('id-ID')} baris. Ekspor Excel tetap memuat seluruh data.</p>}
        </>}
        <div className="bc27-export-row"><button className="bc27-primary" disabled={!result.database.length} onClick={exportExcel}>⇩ EXPORT EXCEL · DATABASE + OUTPUT</button><small>File Excel berisi dua sheet: DATABASE dan OUTPUT.</small></div>
      </section>}
    </>}


      {result&&workflowOption==='standard'&&<section className="bc27-card bc27-crosscheck">
        <div className="bc27-section-head"><div><h2>6. Crosscheck CEISA vs IT Inventory</h2><p>Bagian ini tersedia setelah CEISA berhasil diproses. Hanya kategori 1 · Hasil Produksi yang dibandingkan.</p></div><span className="bc27-step">06</span></div>
        <div className="bc27-cross-flow">
          <div className="bc27-cross-step"><b>A · Master Mapping ERP tunggal</b><p>Gunakan satu file master terbaru dengan kolom CUSTOMER, NO STYLE, dan ITEM CODE CHINLI. Setelah dibaca, mapping disimpan di browser ini sehingga tetap tersedia saat halaman dibuka kembali. Mapping diterapkan sesuai customer, bukan lintas customer.</p>
            <label className="bc27-file-input">Pilih / ganti Master Mapping ERP (.xlsx)<input type="file" accept=".xlsx" onChange={e=>{setErpFile(e.target.files?.[0]||null);setCrosscheck(null);setCrossError('')}}/></label>
            {erpFile&&<small>File dipilih: {erpFile.name}</small>}
            {mappings.length>0&&<small className="bc27-good">Master aktif: {mappings.length.toLocaleString('id-ID')} relasi tersimpan di browser ini.</small>}
            <label className="bc27-mapping-mode">Mode mapping<select value={mappingMode} onChange={e=>setMappingMode(e.target.value)}><option value="replace">Replace · Ganti seluruh master</option><option value="update">Update · Tambahkan/perbarui mapping</option></select></label><button className="bc27-secondary" onClick={downloadMappingTemplate}>DOWNLOAD TEMPLATE MAPPING</button><button className="bc27-secondary" disabled={!erpFile||crossBusy} onClick={async()=>{setCrossBusy(true);setCrossError('');try{if(mappingMode==='replace'&&mappings.length&&!window.confirm('Ganti seluruh master mapping aktif dengan isi file yang dipilih?'))return;const parsed=await parseErpMappings([erpFile]);const next=mappingMode==='update'?[...new Map([...mappings,...parsed].map(x=>[[x.customerKey,x.ceisaKey,x.internalKey].join('|'),x])).values()]:parsed;setMappings(next);setCrosscheck(null);setCrossError('');setNotice('Master Mapping ERP berhasil '+(mappingMode==='replace'?'diganti':'diperbarui')+' dan disimpan di browser ini.')}catch(e){setCrossError(e.message||'Gagal membaca master ERP.')}finally{setCrossBusy(false)}}}>BACA & SIMPAN MASTER ERP</button>
            <div className="bc27-manual-mapping"><b>Tambah mapping manual</b><div><input value={manualCustomer} onChange={e=>setManualCustomer(e.target.value)} placeholder="Customer (nama entitas CEISA)"/><input value={manualCeisa} onChange={e=>setManualCeisa(e.target.value)} placeholder="Kode CEISA / NO STYLE"/><input value={manualInternal} onChange={e=>setManualInternal(e.target.value)} placeholder="Kode internal Chinli"/></div><button className="bc27-secondary" disabled={!manualCustomer.trim()||!manualCeisa.trim()||!manualInternal.trim()} onClick={()=>{const row={customer:manualCustomer.trim(),customerKey:manualCustomer.trim().toUpperCase().replace(/[^A-Z0-9]/g,''),ceisaCode:manualCeisa.trim(),ceisaKey:manualCeisa.trim().toUpperCase().replace(/[^A-Z0-9]/g,''),internalCode:manualInternal.trim(),internalKey:manualInternal.trim().toUpperCase().replace(/[^A-Z0-9]/g,'')};setMappings(prev=>[...prev.filter(x=>![x.customerKey,x.ceisaKey,x.internalKey].every((v,i)=>v===[row.customerKey,row.ceisaKey,row.internalKey][i])),row]);setManualCustomer('');setManualCeisa('');setManualInternal('');setCrosscheck(null)}}>TAMBAH MAPPING</button></div>
            {mappings.length>0&&<small className="bc27-good">{mappings.length.toLocaleString('id-ID')} relasi mapping terbaca.</small>}
          </div>
          <div className="bc27-cross-step"><b>B · Upload IT Inventory</b><p>Gunakan file mutasi barang jadi. Sistem membaca sheet BARANG JADI, kolom KODE BARANG, SAT, dan PENGELUARAN.</p>
            <label className="bc27-file-input">Pilih file IT Inventory (.xls / .xlsx)<input type="file" accept=".xls,.xlsx" onChange={e=>{setInventoryFile(e.target.files?.[0]||null);setInventory(null);setCrosscheck(null);setCrossError('')}}/></label>
            {inventoryFile&&<small>{inventoryFile.name}</small>}
          </div>
        </div>
        <div className="bc27-actions"><button className="bc27-primary" disabled={!inventoryFile||!mappings.length||crossBusy} onClick={async()=>{setCrossBusy(true);setCrossError('');try{const inv=await parseFinishedGoodsInventory(inventoryFile);const rows=buildCrosscheck(result.database,inv.rows,mappings);setInventory(inv);setCrosscheck(rows);setCrossFilter('ALL')}catch(e){setCrossError(e.message||'Crosscheck gagal diproses.')}finally{setCrossBusy(false)}}}>{crossBusy?'MEMPROSES CROSSCHECK...':'JALANKAN CROSSCHECK'}</button><button className="bc27-secondary" disabled={!crosscheck?.length} onClick={()=>exportCrosscheck(crosscheck,period.month+'-'+period.year,inventory?.period)}>EXPORT HASIL EXCEL</button></div>
        {crossError&&<div className="bc27-error">{crossError}</div>}
        {crosscheck&&<>
          <div className="bc27-cross-meta"><span>Periode CEISA: <b>{monthNames[Number(period.month)-1]} {period.year}</b></span><span>Periode IT Inventory: <b>{inventory?.period||'Tidak terdeteksi'}</b></span><span>Baris hasil: <b>{crosscheck.length.toLocaleString('id-ID')}</b></span></div>
          {inventory?.period&&!inventory.period.includes(period.year+'-'+period.month)&&!inventory.period.includes(period.year+'/'+Number(period.month))&&<div className="bc27-warning">Periode berbeda: IT Inventory ({inventory.period}) perlu diperiksa terhadap periode CEISA ({monthNames[Number(period.month)-1]} {period.year}).</div>}
          <div className="bc27-cross-summary">{['MATCH','NOT MATCH','CEISA ONLY','IT ONLY','BELUM ADA MAPPING','MAPPING AMBIGU'].map(s=><button key={s} className={crossFilter===s?'active':''} onClick={()=>setCrossFilter(crossFilter===s?'ALL':s)}><small>{s}</small><b>{crosscheck.filter(r=>r.status===s).length.toLocaleString('id-ID')}</b></button>)}</div>
          <div className="bc27-cross-filter"><label>Filter status<select value={crossFilter} onChange={e=>setCrossFilter(e.target.value)}><option value="ALL">Semua status</option>{['MATCH','NOT MATCH','CEISA ONLY','IT ONLY','BELUM ADA MAPPING','MAPPING AMBIGU'].map(s=><option key={s}>{s}</option>)}</select></label><small>Selisih = Qty CEISA dikurangi Qty IT Inventory.</small></div>
          <div className="bc27-table-wrap"><table className="bc27-table"><thead><tr><th>Status</th><th>Customer</th><th>Kode CEISA</th><th>Uraian Barang</th><th>Kode Internal ERP</th><th>Satuan</th><th>Qty CEISA</th><th>Qty IT Inventory</th><th>Selisih</th><th>Catatan</th></tr></thead><tbody>{crosscheck.filter(r=>crossFilter==='ALL'||r.status===crossFilter).slice(0,1000).map((r,i)=><tr key={r.customer+'-'+r.code+'-'+r.unit+'-'+i}><td><span className={'bc27-status bc27-status-'+r.status.toLowerCase().replace(/[^a-z]+/g,'-')}>{r.status}</span></td><td>{r.customer}</td><td>{r.code}</td><td className="wide-cell">{r.description}</td><td>{(r.internalCodes||[]).join(' + ')}</td><td>{r.unit}</td><td>{fmt(r.quantity)}</td><td>{r.inventoryQuantity==null?'—':fmt(r.inventoryQuantity)}</td><td>{r.difference==null?'—':fmt(r.difference)}</td><td className="wide-cell">{r.note}</td></tr>)}</tbody></table></div>
          {crosscheck.filter(r=>crossFilter==='ALL'||r.status===crossFilter).length>1000&&<p className="bc27-preview-note">Pratinjau dibatasi 1.000 baris. Ekspor Excel memuat seluruh hasil.</p>}
        </>}
      </section>}

    <section className="bc27-card bc27-rules">
      <h2>Aturan pembacaan dan penggabungan</h2>
      <ul>
        <li>Hanya dokumen berkode 27 dengan entitas kode 3 bernama persis <b>CHINLI PLASTIC MATERIALS INDONESIA</b> yang diproses sebagai BC 2.7 Keluar.</li>
        <li>Entitas tujuan dibaca dari kode entitas 8.</li>
        <li>Nomor AJU, nomor daftar, dan tanggal daftar berasal dari HEADER kolom A, CP, dan CQ.</li>
        <li>Kode barang, HS Code, uraian, satuan, jumlah, kategori, dan Spesifikasi Lain berasal dari BARANG kolom D, C, E, J, K, AW, dan I.</li>
        <li>Duplikat Nomor AJU antarfile hanya dihitung satu kali; file pertama yang diproses menjadi sumber data.</li>
        <li>OUTPUT untuk Adonia menggunakan Spesifikasi Lain (BARANG kolom I); perusahaan lain menggunakan Kode Barang.</li>
        <li>Penggabungan OUTPUT: identitas barang + Jenis Satuan + Kategori Barang. Identitas barang Adonia adalah Spesifikasi Lain.</li>
      </ul>
      <div className="bc27-category"><b>Master kategori CEISA</b><div>{Object.entries(CATEGORY_MASTER).map(([code,name])=><span key={code}><b>{code}</b> {name}</span>)}</div></div>
    </section>
  </main>;
}
