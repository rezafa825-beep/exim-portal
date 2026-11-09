const KEY='exim-web-customers-v1';
const DELETED_KEY='exim-web-deleted-customers-v1';
const norm=n=>String(n||'').toUpperCase().replace(/^PT\.?\s*/,'').replace(/\s+/g,' ').trim();
const aliases={INVOICE_NUMBER:['INVOICE NO','INVOICE NUMBER','INVOICE'],ITEM_CODE:['SKU','STYLE','MATERIAL CODE','PRODUCT CODE','ARTICLE'],ITEM_NAME:['ITEM NAME','DESCRIPTION','DESC'],QUANTITY:['QTY','QUANTITY','JUMLAH','TOTAL QTY'],AMOUNT:['AMOUNT'],SURAT_JALAN:['SURAT','SURAT JALAN'],GROSS_WEIGHT:['GW','GROSS WT','GROSS WEIGHT'],NET_WEIGHT:['NW','NET WT','NET WEIGHT'],PACKAGE:['KEMASAN','PACKAGE','PACKAGE TYPE']};
const make=(name,invoice,pl)=>({id:name,status:'ACTIVE',customer_name:name,mappings:{invoice:Object.fromEntries(Object.entries(invoice).map(([k,v])=>[k,{primary:v,alternatives:(aliases[k]||[]).filter(a=>a!==v)}])),packing_list:Object.fromEntries(Object.entries(pl).map(([k,v])=>[k,{primary:v,alternatives:(aliases[k]||[]).filter(a=>a!==v)}]))},row_rule:{item_identifier:'ITEM NAME',aggregation_labels:'TOTAL;SUBTOTAL;GRAND TOTAL',continue_after_aggregation:true}});
const defaults=[make('PT. SHOETOWN LIGUNG INDONESIA',{INVOICE_NUMBER:'INVOICE NO',ITEM_CODE:'STYLE',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',AMOUNT:'AMOUNT',SURAT_JALAN:'SURAT'},{PACKING_LIST_NO:'INVOICE NO',ITEM_CODE:'STYLE',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',GROSS_WEIGHT:'GW',NET_WEIGHT:'NW',PACKAGE:'KEMASAN'}),make('PT. YIH QUAN FOOTWEAR INDONESIA',{INVOICE_NUMBER:'INVOICE NO',ITEM_CODE:'SKU',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',AMOUNT:'AMOUNT',SURAT_JALAN:'SURAT'},{PACKING_LIST_NO:'INVOICE NO',ITEM_CODE:'SKU',ITEM_NAME:'ITEM NAME',QUANTITY:'TOTAL QTY',GROSS_WEIGHT:'GW',NET_WEIGHT:'NW',PACKAGE:'KEMASAN'}),make('PT. ADONIA FOOTWEAR INDONESIA',{INVOICE_NUMBER:'INVOICE NO',ITEM_CODE:'MATERIAL CODE',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',AMOUNT:'AMOUNT',SURAT_JALAN:'SURAT JALAN'},{PACKING_LIST_NO:'INVOICE NO',ITEM_CODE:'MATERIAL CODE',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',GROSS_WEIGHT:'GW',NET_WEIGHT:'NW',PACKAGE:'KEMASAN'}),make('PT. CHINLI PLASTIC TECHNOLOGY INDONESIA',{INVOICE_NUMBER:'INVOICE NO',ITEM_CODE:'MATERIAL CODE CUSTOMER',ITEM_NAME:'ITEM NAME',QUANTITY:'QTY',AMOUNT:'AMOUNT',SURAT_JALAN:'SURAT JALAN'},{PACKING_LIST_NO:'INVOICE NO',ITEM_CODE:'ITEM CODE',ITEM_NAME:'ITEM NAME',QUANTITY:'TOTAL QTY (PAIRS)',GROSS_WEIGHT:'GW',NET_WEIGHT:'NW',PACKAGE:'KEMASAN'})];
defaults[3].default_currency='USD';
function readDeleted(){try{return new Set((JSON.parse(localStorage.getItem(DELETED_KEY)||'[]')||[]).map(norm));}catch{return new Set();}}
export function saveCustomers(rows){localStorage.setItem(KEY,JSON.stringify(rows));}
export function getCustomers(){
  try {
    const stored=JSON.parse(localStorage.getItem(KEY)||'[]');
    const deleted=readDeleted();
    const merged=[...(Array.isArray(stored)?stored:[])];
    for(const d of defaults){
      if(deleted.has(norm(d.customer_name))) continue;
      const existing=merged.find(c=>norm(c.customer_name)===norm(d.customer_name));
      if(!existing){merged.push(structuredClone(d));continue;}
      if(d.default_currency&&!existing.default_currency) existing.default_currency=d.default_currency;
      for(const doc of Object.keys(d.mappings||{})){
        existing.mappings=existing.mappings||{}; existing.mappings[doc]=existing.mappings[doc]||{};
        for(const [field,defaultMap] of Object.entries(d.mappings[doc]||{})){
          const current=existing.mappings[doc][field];
          if(!current||!String(current.primary||'').trim()) existing.mappings[doc][field]={...defaultMap,...(current||{}),primary:defaultMap.primary};
        }
      }
      if(!existing.row_rule) existing.row_rule=d.row_rule;
    }
    // Hapus duplikat berdasarkan nama ternormalisasi. Prioritaskan record ACTIVE,
    // lalu record dengan mapping paling lengkap agar konfigurasi yang terpakai konsisten.
    const unique=new Map(); const duplicateNames=new Set();
    for(const row of merged){
      if(!row?.customer_name||deleted.has(norm(row.customer_name))) continue;
      const key=norm(row.customer_name); const old=unique.get(key); if(old) duplicateNames.add(key);
      const countMappings=x=>Object.values(x.mappings||{}).reduce((n,doc)=>n+Object.values(doc||{}).filter(v=>String(v?.primary||'').trim()).length,0);
      if(!old || (String(row.status).toUpperCase()==='ACTIVE'&&String(old.status).toUpperCase()!=='ACTIVE') || (String(row.status).toUpperCase()===String(old.status).toUpperCase()&&countMappings(row)>countMappings(old))) unique.set(key,row);
    }
    const rows=[...unique.values()];
    // Jika data Chinli lama terduplikasi dan semuanya nonaktif, pulihkan satu record aktif.
    // Setelah duplikat dibersihkan, perubahan status manual tetap dihormati.
    const chinliKey=norm('PT. CHINLI PLASTIC TECHNOLOGY INDONESIA');
    const chinli=rows.find(c=>norm(c.customer_name)===chinliKey);
    if(chinli&&duplicateNames.has(chinliKey)) chinli.status='ACTIVE';
    localStorage.setItem(KEY,JSON.stringify(rows));
    return rows;
  } catch { return defaults.map(x=>structuredClone(x)); }
}
export function findCustomer(name){
  const n=norm(name); if(!n) return null;
  const rows=getCustomers().filter(c=>String(c.status||'ACTIVE').toUpperCase()==='ACTIVE');
  const exact=rows.find(c=>norm(c.customer_name)===n); if(exact) return exact;
  const scored=rows.map(c=>{const x=norm(c.customer_name);return {c,score:x===n?100:(x.includes(n)||n.includes(x)?50:0)};}).filter(x=>x.score>0).sort((a,b)=>b.score-a.score);
  return scored[0]?.c||null;
}
export function upsertCustomer(row){
  const rows=getCustomers(); const key=norm(row.customer_name);
  if(!key) throw new Error('Nama customer wajib diisi.');
  const duplicate=rows.find(x=>norm(x.customer_name)===key&&x.id!==row.id);
  if(duplicate) throw new Error(`Customer dengan nama "${duplicate.customer_name}" sudah terdaftar. Gunakan data yang sudah ada.`);
  const i=rows.findIndex(x=>x.id===row.id); if(i>=0) rows[i]={...row,customer_name:String(row.customer_name).trim()}; else rows.push({...row,customer_name:String(row.customer_name).trim(),id:row.id||crypto.randomUUID()});
  const deleted=readDeleted(); deleted.delete(key); localStorage.setItem(DELETED_KEY,JSON.stringify([...deleted]));
  saveCustomers(rows); return rows;
}
export function removeCustomer(id){
  const rows=getCustomers(); const target=rows.find(x=>x.id===id); if(!target) return rows;
  const deleted=readDeleted(); deleted.add(norm(target.customer_name)); localStorage.setItem(DELETED_KEY,JSON.stringify([...deleted]));
  const next=rows.filter(x=>x.id!==id); saveCustomers(next); return next;
}
