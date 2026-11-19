import * as XLSX from 'xlsx';

const clean = v => String(v ?? '').replace(/\u00a0/g, ' ').replace(/[\r\n]+/g, ' ').trim();
const norm = v => clean(v).toUpperCase().replace(/[^A-Z0-9]/g, '');
const rowsOf = (wb, name, fileName) => {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(fileName + ': sheet ' + name + ' tidak ditemukan.');
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
};
const findHeader = (rows, required, max = 50) => {
  for (let i = 0; i < Math.min(rows.length, max); i++) {
    const h = (rows[i] || []).map(norm);
    if (required.every(key => h.some(v => v === key || v.includes(key)))) return i;
  }
  return -1;
};
const findColumn = (header, key) => header.findIndex(v => norm(v) === key || norm(v).includes(key));
const asKey = v => clean(v).toUpperCase().replace(/\s+/g, '');

function parseInvoiceWorkbook(file) {
  const wb = XLSX.read(file.arrayBuffer ? undefined : undefined);
  return wb;
}

export async function buildAdoniaSkuOutputs(ceisaFiles, invoiceFiles) {
  if (!ceisaFiles?.length) throw new Error('File CEISA belum tersedia.');
  if (!invoiceFiles?.length) throw new Error('Pilih minimal satu file Excel Invoice Adonia.');
  const invoiceItems = new Map();
  const invoiceMeta = new Map();
  const invoiceErrors = [];
  let invoiceSheetsRead = 0;

  for (const file of invoiceFiles) {
    if (!/\.xlsx?$/i.test(file.name)) { invoiceErrors.push(file.name + ': format tidak didukung.'); continue; }
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', raw: true, blankrows: false });
        const hi = findHeader(rows, ['NO', 'SKU', 'MATERIALCODE'], 60);
        if (hi < 0) continue;
        const header = (rows[hi] || []).map(norm);
        const noCol = findColumn(header, 'NO');
        const skuCol = findColumn(header, 'SKU');
        const materialCol = findColumn(header, 'MATERIALCODE');
        if ([noCol, skuCol, materialCol].some(x => x < 0)) continue;
        const topText = rows.slice(0, hi + 1).flat().map(clean).join(' | ');
        const invMatch = topText.match(/INVOICE\s*NO\s*:?\s*([A-Z0-9][A-Z0-9./_-]*)/i);
        const shipToAdonia = /ADONIA/i.test(topText);
        if (!invMatch || !shipToAdonia) continue;
        const invoiceNo = clean(invMatch[1]).toUpperCase();
        invoiceSheetsRead++;
        invoiceMeta.set(invoiceNo, { invoiceNo, file: file.name, sheet: sheetName });
        for (const row of rows.slice(hi + 1)) {
          const lineNo = clean(row[noCol]);
          const sku = clean(row[skuCol]);
          const materialCode = clean(row[materialCol]);
          if (!/^\d+$/.test(lineNo) || !sku || !materialCode) continue;
          const key = invoiceNo + '|' + String(Number(lineNo));
          const item = { invoiceNo, lineNo: String(Number(lineNo)), sku, materialCode, materialKey: norm(materialCode), source: file.name, sheet: sheetName };
          if (!invoiceItems.has(key)) invoiceItems.set(key, []);
          invoiceItems.get(key).push(item);
        }
      }
    } catch (e) {
      invoiceErrors.push(file.name + ': ' + (e.message || 'gagal dibaca.'));
    }
  }
  if (!invoiceSheetsRead || !invoiceItems.size) {
    throw new Error(invoiceErrors.join('\n') || 'Tidak ditemukan sheet Invoice Adonia dengan kolom NO, SKU, MATERIAL CODE dan nomor invoice.');
  }

  const outputs = [];
  const details = [];
  const issues = [...invoiceErrors];
  let adoniaAjuCount = 0;
  let itemMatched = 0;
  let itemUnmatched = 0;
  let itemSkipped = 0;

  for (const file of ceisaFiles) {
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellStyles: true });
      const headerRows = rowsOf(wb, 'HEADER', file.name);
      const entityRows = rowsOf(wb, 'ENTITAS', file.name);
      const documentRows = rowsOf(wb, 'DOKUMEN', file.name);
      const goodsRows = rowsOf(wb, 'BARANG', file.name);
      const headerData = headerRows.slice(1);
      const entitiesByAju = new Map();
      for (const row of entityRows.slice(1)) {
        const aju = clean(row[0]);
        if (!aju) continue;
        if (!entitiesByAju.has(aju)) entitiesByAju.set(aju, []);
        entitiesByAju.get(aju).push(row);
      }
      const docsByAju = new Map();
      for (const row of documentRows.slice(1)) {
        const aju = clean(row[0]);
        if (!aju) continue;
        if (!docsByAju.has(aju)) docsByAju.set(aju, []);
        docsByAju.get(aju).push(row);
      }
      const ajuInvoice = new Map();
      for (const h of headerData) {
        const aju = clean(h[0]);
        if (!aju || clean(h[1]) !== '27') continue;
        const entities = entitiesByAju.get(aju) || [];
        const chinli = entities.some(r => clean(r[2]) === '3' && norm(r[5]) === 'CHINLIPLASTICMATERIALSINDONESIA');
        const adonia = entities.some(r => clean(r[2]) === '8' && /ADONIA/i.test(clean(r[5])));
        if (!chinli || !adonia) continue;
        const docs = docsByAju.get(aju) || [];
        const invoiceDoc = docs.find(r => clean(r[2]) === '380') || docs.find(r => clean(r[2]) === '217');
        const invoiceNo = clean(invoiceDoc?.[3]).toUpperCase();
        if (!invoiceNo) {
          issues.push(file.name + ' · ' + aju + ': nomor invoice tidak ditemukan pada DOKUMEN (kode 380/217).');
          continue;
        }
        ajuInvoice.set(aju, invoiceNo);
        adoniaAjuCount++;
      }

      let changed = 0;
      const goodsSheet = wb.Sheets.BARANG;
      if (!goodsSheet) throw new Error('Sheet BARANG tidak ditemukan.');
      for (let r = 1; r < goodsRows.length; r++) {
        const row = goodsRows[r] || [];
        const aju = clean(row[0]);
        const invoiceNo = ajuInvoice.get(aju);
        if (!invoiceNo) continue;
        const lineNo = clean(row[1]);
        const materialCode = clean(row[3]);
        const key = invoiceNo + '|' + String(Number(lineNo));
        const candidates = invoiceItems.get(key) || [];
        if (!candidates.length) {
          itemUnmatched++;
          details.push({ file: file.name, aju, invoiceNo, lineNo, materialCode, sku: '', status: 'INVOICE ITEM NOT FOUND', note: 'Nomor urut item tidak ditemukan pada Invoice.' });
          continue;
        }
        const valid = candidates.filter(x => x.materialKey === norm(materialCode));
        if (valid.length !== 1) {
          itemUnmatched++;
          details.push({ file: file.name, aju, invoiceNo, lineNo, materialCode, sku: candidates.map(x => x.sku).join(' / '), status: valid.length > 1 ? 'DUPLICATE INVOICE ITEM' : 'MATERIAL CODE MISMATCH', note: valid.length > 1 ? 'Lebih dari satu baris Invoice cocok.' : 'MATERIAL CODE Invoice berbeda dari KODE BARANG CEISA.' });
          continue;
        }
        const sku = valid[0].sku;
        const existing = clean(row[8]);
        if (existing && existing !== '-' && existing !== sku) {
          itemSkipped++;
          details.push({ file: file.name, aju, invoiceNo, lineNo, materialCode, sku, status: 'SPECIFICATION CONFLICT', note: 'Spesifikasi Lain sudah berisi nilai berbeda; tidak ditimpa.' });
          continue;
        }
        const cellAddress = XLSX.utils.encode_cell({ r, c: 8 });
        if (!goodsSheet[cellAddress]) goodsSheet[cellAddress] = { t: 's', v: sku };
        else { goodsSheet[cellAddress].t = 's'; goodsSheet[cellAddress].v = sku; delete goodsSheet[cellAddress].w; }
        changed++;
        itemMatched++;
        details.push({ file: file.name, aju, invoiceNo, lineNo, materialCode, sku, status: 'MATCHED', note: 'SKU berhasil dimasukkan ke Spesifikasi Lain.' });
      }
      outputs.push({
        name: file.name.replace(/\.xlsx$/i, '') + '_ADD_SKU_ADONIA.xlsx',
        data: XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true }),
        changed
      });
    } catch (e) {
      issues.push(file.name + ': ' + (e.message || 'gagal diproses.'));
    }
  }

  if (!outputs.length) throw new Error(issues.join('\n') || 'Tidak ada file CEISA yang dapat diproses.');
  return {
    outputs, details, issues,
    summary: { invoiceFiles: invoiceFiles.length, invoiceSheetsRead, adoniaAjuCount, itemMatched, itemUnmatched, itemSkipped, outputFiles: outputs.length }
  };
}
