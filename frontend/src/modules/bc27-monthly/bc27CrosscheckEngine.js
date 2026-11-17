import * as XLSX from 'xlsx';

const clean = v => String(v ?? '').replace(/\u00a0/g, ' ').trim();
const norm = v => clean(v).toUpperCase().replace(/[^A-Z0-9]/g, '');
const number = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  let value = clean(v).replace(/\s/g, '');
  if (!value) return 0;
  if (value.includes(',') && value.includes('.')) {
    value = value.lastIndexOf(',') > value.lastIndexOf('.')
      ? value.replace(/\./g, '').replace(',', '.')
      : value.replace(/,/g, '');
  } else if (value.includes(',')) {
    value = value.replace(',', '.');
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};
const rowsOf = (wb,name) => { const ws=wb.Sheets[name]; if(!ws) throw new Error('Sheet '+name+' tidak ditemukan.'); return XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:false}); };
const findHeader = (rows,keys,max=40) => { for(let i=0;i<Math.min(rows.length,max);i++){const h=(rows[i]||[]).map(norm);if(keys.every(k=>h.some(v=>v===k||v.includes(k))))return i;} return -1; };
const aggregate = (rows,keyFn) => { const map=new Map(); for(const r of rows){const k=keyFn(r);if(!map.has(k))map.set(k,{...r,quantity:0});map.get(k).quantity+=r.quantity;} return [...map.values()]; };
const dateText = v => { if(v instanceof Date&&!Number.isNaN(v.getTime()))return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0'); if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);if(d)return d.y+'-'+String(d.m).padStart(2,'0')+'-'+String(d.d).padStart(2,'0');} const s=clean(v);let m=s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);if(m)return m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');m=s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);return m?m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'):''; };

export async function parseErpMappings(files) {
  const out=[],errors=[];
  for(const file of files||[]) {
    try {
      const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
      for(const sheetName of wb.SheetNames) {
        const rows=rowsOf(wb,sheetName), hi=findHeader(rows,['CUSTOMER','NOSTYLE','ITEMCODECHINLI'],35);
        if(hi<0) continue;
        const h=rows[hi].map(norm), col=k=>h.findIndex(v=>v===k||v.includes(k));
        const c=col('CUSTOMER'), s=col('NOSTYLE'), i=col('ITEMCODECHINLI');
        for(const row of rows.slice(hi+1)) {
          const customer=clean(row[c]),ceisaCode=clean(row[s]),internalCode=clean(row[i]);
          if(customer&&ceisaCode&&internalCode) out.push({customer,customerKey:norm(customer),ceisaCode,ceisaKey:norm(ceisaCode),internalCode,internalKey:norm(internalCode),source:file.name,sheet:sheetName});
        }
      }
    } catch(e) { errors.push(file.name+': '+(e.message||'gagal dibaca')); }
  }
  if(!out.length) throw new Error(errors.length?errors.join('\n'):'Tidak ditemukan sheet dengan header CUSTOMER, NO STYLE, dan ITEM CODE CHINLI.');
  const unique=new Map(); out.forEach(r=>unique.set([r.customerKey,r.ceisaKey,r.internalKey].join('|'),r)); return [...unique.values()];
}

export async function parseFinishedGoodsInventory(file) {
  if(!file) throw new Error('Pilih file IT Inventory terlebih dahulu.');
  if(!/\.xlsx?$/i.test(file.name)) throw new Error('Format file IT Inventory harus .xls atau .xlsx.');
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
  const sheetName=wb.SheetNames.find(n=>norm(n)==='BARANGJADI'); if(!sheetName) throw new Error(file.name+': sheet BARANG JADI tidak ditemukan.'); const rows=rowsOf(wb,sheetName);
  const hi=findHeader(rows,['KODEBARANG','PENGELUARAN']);
  if(hi<0) throw new Error(file.name+': header KODE BARANG, SAT, dan PENGELUARAN tidak ditemukan.');
  const h=rows[hi].map(norm), col=k=>h.findIndex(v=>v===k||v.includes(k));
  const c=col('KODEBARANG'), n=col('NAMABARANG'), u=col('SAT'), q=col('PENGELUARAN');
  if(c<0||u<0||q<0) throw new Error('Kolom KODE BARANG, SAT, atau PENGELUARAN tidak lengkap.');
  const raw=[];
  for(const row of rows.slice(hi+1)){const code=clean(row[c]),unit=clean(row[u]);if(!code||!unit)continue;raw.push({internalCode:code,internalKey:norm(code),description:n>=0?clean(row[n]):'',unit,unitKey:norm(unit),quantity:number(row[q])});}
  if(!raw.length) throw new Error('Tidak ditemukan data barang pada sheet '+sheetName+'.');
  const periodLine=rows.slice(0,hi).flat().map(clean).find(v=>/\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(v)||/\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(v))||'';
  const dates=[...periodLine.matchAll(/(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})/g)].map(m=>dateText(m[1]));
  return {rows:aggregate(raw,r=>r.internalKey+'|'+r.unitKey),period:dates.join(' s/d ')||periodLine,sourceSheet:sheetName,sourceRows:raw.length};
}

export function buildCrosscheck(database,inventoryRows,mappings) {
  const ceisaRaw=(database||[]).filter(r=>clean(r['Kategori Barang']).toUpperCase()==='HASIL PRODUKSI').map(r=>{
    const customer=clean(r['Entitas Perusahaan Tujuan']), adonia=/ADONIA/i.test(customer), code=clean(adonia?r['Spesifikasi Lain']:r['Kode Barang']);
    return {customer,customerKey:norm(customer),code,codeKey:norm(code),description:clean(r['Uraian Barang']),unit:clean(r['Jenis Satuan']),unitKey:norm(r['Jenis Satuan']),quantity:number(r['Jumlah Satuan'])};
  }).filter(r=>r.customerKey&&r.codeKey&&r.unitKey);
  if(!ceisaRaw.length) throw new Error('Tidak ditemukan barang kategori 1 · Hasil Produksi pada hasil CEISA.');
  const ceisa=aggregate(ceisaRaw,r=>[r.customerKey,r.codeKey,r.unitKey].join('|'));
  const ceisaByKey=new Map(ceisa.map(r=>[[r.customerKey,r.codeKey,r.unitKey].join('|'),r]));
  const mapByCeisa=new Map(), mapByInternal=new Map();
  for(const m of mappings||[]) {
    const a=[m.customerKey||norm(m.customer),m.ceisaKey||norm(m.ceisaCode)], b=[m.customerKey||norm(m.customer),m.internalKey||norm(m.internalCode)];
    const ck=a.join('|'),ik=b.join('|');
    if(!mapByCeisa.has(ck))mapByCeisa.set(ck,new Map()); mapByCeisa.get(ck).set(m.internalKey||norm(m.internalCode),m);
    if(!mapByInternal.has(m.internalKey||norm(m.internalCode)))mapByInternal.set(m.internalKey||norm(m.internalCode),new Map());
    mapByInternal.get(m.internalKey||norm(m.internalCode)).set(ck,m);
  }
  const itMap=new Map((inventoryRows||[]).map(r=>[[r.internalKey||norm(r.internalCode),r.unitKey||norm(r.unit)].join('|'),r]));
  const used=new Set(), rows=[];
  for(const c of ceisa) {
    const key=[c.customerKey,c.codeKey].join('|'), internals=[...(mapByCeisa.get(key)?.values()||[])];
    if(!internals.length){rows.push({...c,internalCodes:[],inventoryQuantity:null,difference:null,status:'BELUM ADA MAPPING',note:'Kode CEISA belum ada di master ERP untuk customer ini.'});continue;}
    const ambiguous=internals.some(m=>(mapByInternal.get(m.internalKey||norm(m.internalCode))?.size||0)>1);
    if(ambiguous){internals.forEach(m=>used.add([m.internalKey||norm(m.internalCode),c.unitKey].join('|')));rows.push({...c,internalCodes:internals.map(m=>m.internalCode),inventoryQuantity:null,difference:null,status:'MAPPING AMBIGU',note:'Satu kode internal terhubung ke lebih dari satu customer/kode CEISA pada master ERP. Perlu verifikasi manual.'});continue;}
    const found=internals.map(m=>itMap.get([m.internalKey||norm(m.internalCode),c.unitKey].join('|')));
    const present=found.filter(Boolean); present.forEach(r=>used.add([r.internalKey||norm(r.internalCode),r.unitKey||norm(r.unit)].join('|')));
    const invQty=present.reduce((s,r)=>s+r.quantity,0), missing=internals.filter((m,i)=>!found[i]), diff=c.quantity-invQty;
    const status=missing.length?(present.length?'NOT MATCH':'CEISA ONLY'):(Math.abs(diff)<1e-8?'MATCH':'NOT MATCH');
    rows.push({...c,internalCodes:internals.map(m=>m.internalCode),inventoryQuantity:present.length?invQty:0,difference:present.length?diff:null,status,note:missing.length?'Tidak ditemukan di IT Inventory: '+missing.map(m=>m.internalCode).join(', '):(status==='MATCH'?'Jumlah CEISA dan IT Inventory sesuai.':'Jumlah CEISA dan IT Inventory berbeda.')});
  }
  for(const it of inventoryRows||[]) {
    const ik=it.internalKey||norm(it.internalCode), unit=it.unitKey||norm(it.unit), key=[ik,unit].join('|');
    if(used.has(key)||Math.abs(it.quantity)<1e-12)continue;
    const refs=[...(mapByInternal.get(ik)?.values()||[])];
    const ambiguous=refs.length>1;
    rows.push({customer:[...new Set(refs.map(x=>x.customer))].join(', ')||'-',code:it.internalCode,description:it.description,unit:it.unit,quantity:0,inventoryQuantity:it.quantity,difference:-it.quantity,internalCodes:[it.internalCode],status:ambiguous?'MAPPING AMBIGU':refs.length?'IT ONLY':'BELUM ADA MAPPING',note:ambiguous?'Kode internal mempunyai lebih dari satu pasangan customer/kode CEISA pada master ERP.':refs.length?'Pengeluaran IT Inventory tidak ditemukan pada pasangan CEISA untuk periode ini.':'Kode internal tidak ditemukan pada master ERP.'});
  }
  return rows.sort((a,b)=>a.status.localeCompare(b.status)||a.customer.localeCompare(b.customer)||a.code.localeCompare(b.code));
}

export function downloadMappingTemplate() { const wb=XLSX.utils.book_new(); const ws=XLSX.utils.aoa_to_sheet([['CUSTOMER','NO STYLE','ITEM CODE CHINLI'],['','','']]); ws['!cols']=[36,28,30].map(wch=>({wch})); XLSX.utils.book_append_sheet(wb,ws,'MAPPING'); XLSX.writeFile(wb,'Template_Mapping_ERP_BC27.xlsx'); }

export function exportCrosscheck(rows,period,inventoryPeriod) {
  const wb=XLSX.utils.book_new(), data=rows.map(r=>({'Status':r.status,'Customer':r.customer,'Kode CEISA / Internal':r.code,'Uraian Barang':r.description,'Kode Internal ERP':(r.internalCodes||[]).join(' + '),'Satuan':r.unit,'Qty CEISA':r.quantity,'Qty IT Inventory':r.inventoryQuantity??'','Selisih (CEISA - IT)':r.difference??'','Catatan':r.note}));
  const ws=XLSX.utils.json_to_sheet(data); ws['!cols']=[20,34,28,48,40,14,18,20,22,72].map(wch=>({wch})); if(ws['!ref'])ws['!autofilter']={ref:ws['!ref']}; XLSX.utils.book_append_sheet(wb,ws,'CROSSCHECK');
  const summary=XLSX.utils.aoa_to_sheet([['Keterangan','Nilai'],['Periode CEISA',period],['Periode IT Inventory',inventoryPeriod||'Tidak terdeteksi'],['Total baris',rows.length],...['MATCH','NOT MATCH','CEISA ONLY','IT ONLY','BELUM ADA MAPPING','MAPPING AMBIGU'].map(s=>[s,rows.filter(r=>r.status===s).length])]); XLSX.utils.book_append_sheet(wb,summary,'RINGKASAN');
  XLSX.writeFile(wb,'CROSSCHECK_IT_CEISA_'+period+'_'+Date.now()+'.xlsx');
}
