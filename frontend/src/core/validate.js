import { readExcelFile, detectWorkbook } from './excelReader';
import { parseInvoice, parsePackingList, parseDraft, parseSuratJalan } from './parsers';
import { crosscheck } from './crosscheck';
import { findCustomer } from './customers';

// Jika satu/lebih file berisi beberapa set dokumen (beberapa sheet Invoice/Packing List/Surat Jalan),
// pilih sheet yang nomornya sama dengan Draft. Bila tidak ada yang cocok, pakai yang pertama (perilaku sebelumnya).
const normNo=v=>String(v||'').replace(/\s+/g,' ').trim().toUpperCase();
export function selectDocuments(docs,customer,draft,currencyPreference='AUTO'){
  const notes=[];
  const pick=(type,label,draftNo,readNo)=>{
    const cands=docs.filter(x=>x.type===type);
    if(cands.length<=1||!normNo(draftNo)) return cands[0];
    const hit=cands.find(c=>{try{return normNo(readNo(c))===normNo(draftNo);}catch{return false;}});
    const used=hit||cands[0];
    notes.push(hit
      ? `${label}: terdapat ${cands.length} sheet; dipakai sheet "${used.sheetName}" karena nomornya sama dengan Draft (${draftNo}).`
      : `${label}: terdapat ${cands.length} sheet tetapi tidak ada yang nomornya sama dengan Draft (${draftNo}); dipakai sheet pertama "${used.sheetName}".`);
    return used;
  };
  const invRec=pick('invoice','Invoice',draft.invoice_no,d=>parseInvoice(d.file,customer.mappings.invoice,d.sheetName,currencyPreference).invoice_no);
  const plRec=pick('packing_list','Packing List',draft.packing_list_no,d=>parsePackingList(d.file,customer.mappings.packing_list,customer.row_rule,d.sheetName).packing_list_no);
  const sjRec=pick('surat_jalan','Surat Jalan',draft.surat_jalan,d=>parseSuratJalan(d.file,customer.mappings.surat_jalan||customer.mappings.invoice,d.sheetName).surat_jalan);
  return {invRec,plRec,sjRec,notes};
}

export async function validateFiles(files,currencyPreference='AUTO'){
  const fileData=await Promise.all(files.map(readExcelFile));
  const docs=fileData.flatMap(file=>detectWorkbook(file).map(d=>({...d,file})));
  const detected_documents=docs.map(x=>({file:x.file.name,type:x.type,sheet:x.sheetName})); const draftRec=docs.find(x=>x.type==='draft_exim');
  if(!draftRec) throw new Error('Draft EXIM tidak terdeteksi.');
  const draft=parseDraft(draftRec.file); const customer=findCustomer(draft.company_receiver);
  if(!customer) throw new Error(`Customer belum ada di Customer Master: ${draft.company_receiver||'tidak terbaca'}`);
  // AUTO mengikuti preferensi mata uang customer jika tersedia; pilihan IDR/USD manual tetap menjadi prioritas.
  const effectiveCurrency=String(currencyPreference||'AUTO').toUpperCase()==='AUTO'?(customer.default_currency||'AUTO'):String(currencyPreference).toUpperCase();
  const sel=selectDocuments(docs,customer,draft,effectiveCurrency); const invRec=sel.invRec, plRec=sel.plRec, sjRec=sel.sjRec;
  if(!invRec||!plRec) throw new Error('Invoice dan Packing List wajib tersedia.');
  const inv=parseInvoice(invRec.file,customer.mappings.invoice,invRec.sheetName,effectiveCurrency); const pl=parsePackingList(plRec.file,customer.mappings.packing_list,customer.row_rule,plRec.sheetName); const sj=sjRec ? parseSuratJalan(sjRec.file,customer.mappings.surat_jalan||customer.mappings.invoice,sjRec.sheetName) : null; if(sj?.surat_jalan) inv.surat_jalan=sj.surat_jalan;
  const issues=[];if(!inv.invoice_no)issues.push('Invoice Number tidak terbaca');if(!pl.packing_list_no)issues.push('Packing List Number tidak terbaca');if(!inv.items.length)issues.push('Item Invoice tidak terbaca');if(!pl.items.length)issues.push('Item Packing List tidak terbaca');if(issues.length)throw new Error(`Dokumen belum dapat dipetakan dengan yakin: ${issues.join('; ')}. Periksa struktur/header file.`);
  const result=crosscheck(inv,pl,draft,sj,{currencyPreference:effectiveCurrency,customerName:customer.customer_name}); result.notes=[...(result.notes||[]),...sel.notes];return {session_id:crypto.randomUUID(),customer:customer.customer_name,currency_preference:effectiveCurrency,invoice:inv,packing_list:pl,surat_jalan:sj,draft,result,detected_documents,created_at:new Date().toISOString()};
}
