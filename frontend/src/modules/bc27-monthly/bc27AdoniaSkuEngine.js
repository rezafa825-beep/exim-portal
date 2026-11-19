import * as XLSX from 'xlsx';

export const ADONIA_OUTPUT_HEADERS = [
  'Nomor AJU',
  'Nomor Daftar',
  'Tanggal Daftar',
  'Kode Barang',
  'Uraian Barang',
  'Spesifikasi Lain',
  'Jenis Satuan',
  'Kuantiti'
];

const clean = value => String(value ?? '').replace(/\\u00a0/g, ' ').replace(/[\\r\\n]+/g, ' ').trim();
const norm = value => clean(value).toUpperCase().replace(/[^A-Z0-9]/g, '');
const rowsOf = (workbook, name, fileName) => {
  const sheet = workbook.Sheets[name];
  if (!sheet) throw new Error(fileName + ': sheet ' + name + ' tidak ditemukan.');
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: true });
};
const findHeader = (rows, required, max = 60) => {
  for (let i = 0; i < Math.min(rows.length, max); i++) {
    const header = (rows[i] || []).map(norm);
    if (required.every(key => header.includes(key))) return i;
  }
  return -1;
};
const findColumn = (header, key) => header.findIndex(value => norm(value) === key);

function formatCeisaDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getFullYear() + '-' + String(value.getMonth() + 1).padStart(2, '0') + '-' + String(value.getDate()).padStart(2, '0');
  }
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return parsed.y + '-' + String(parsed.m).padStart(2, '0') + '-' + String(parsed.d).padStart(2, '0');
  }
  const raw = clean(value);
  if (!raw) return '';
  let match = raw.match(/^(\\d{4})[-/](\\d{1,2})[-/](\\d{1,2})/);
  if (match) return match[1] + '-' + String(match[2]).padStart(2, '0') + '-' + String(match[3]).padStart(2, '0');
  match = raw.match(/^(\\d{1,2})[-/](\\d{1,2})[-/](\\d{4})$/);
  if (match) return match[3] + '-' + String(match[2]).padStart(2, '0') + '-' + String(match[1]).padStart(2, '0');
  return raw;
}

function mergeSpecification(existingValue, skuValue) {
  const existing = clean(existingValue);
  const sku = clean(skuValue);
  if (!sku) return existing === '-' ? '' : existing;
  if (!existing || existing === '-') return sku;
  if (norm(existing).includes(norm(sku))) return existing;
  return existing + ' | ' + sku;
}

function numberValue(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const raw = clean(value).replace(/\\s/g, '');
  if (!raw) return 0;
  const parsed = Number(raw.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : value;
}

export async function buildAdoniaSkuOutputs(ceisaFiles, invoiceFiles, period) {
  if (!ceisaFiles?.length) throw new Error('File CEISA belum tersedia.');
  if (!invoiceFiles?.length) throw new Error('Pilih minimal satu file Excel Invoice Adonia.');

  // Index Invoice berdasarkan Nomor Invoice + Nomor Urut Item.
  const invoiceItems = new Map();
  const invoiceErrors = [];
  let invoiceSheetsRead = 0;

  for (const file of invoiceFiles) {
    if (!/\\.xlsx?$/i.test(file.name)) {
      invoiceErrors.push(file.name + ': format tidak didukung.');
      continue;
    }
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      for (const sheetName of workbook.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true, blankrows: false });
        const headerIndex = findHeader(rows, ['NO', 'SKU', 'MATERIALCODE']);
        if (headerIndex < 0) continue;
        const header = (rows[headerIndex] || []).map(norm);
        const noColumn = findColumn(header, 'NO');
        const skuColumn = findColumn(header, 'SKU');
        const materialColumn = findColumn(header, 'MATERIALCODE');
        if ([noColumn, skuColumn, materialColumn].some(index => index < 0)) continue;

        const topText = rows.slice(0, headerIndex + 1).flat().map(clean).join(' | ');
        const invoiceMatch = topText.match(/INVOICE\\s*NO\\s*:?\\s*([A-Z0-9][A-Z0-9./_-]*)/i);
        const isAdoniaInvoice = /ADONIA/i.test(topText);
        if (!invoiceMatch || !isAdoniaInvoice) continue;

        const invoiceNo = clean(invoiceMatch[1]).toUpperCase();
        invoiceSheetsRead++;
        for (const row of rows.slice(headerIndex + 1)) {
          const lineNo = clean(row[noColumn]);
          const sku = clean(row[skuColumn]);
          const materialCode = clean(row[materialColumn]);
          if (!/^\\d+$/.test(lineNo) || !sku || !materialCode) continue;
          const key = invoiceNo + '|' + String(Number(lineNo));
          const item = {
            invoiceNo,
            lineNo: String(Number(lineNo)),
            sku,
            materialKey: norm(materialCode),
            source: file.name,
            sheet: sheetName
          };
          if (!invoiceItems.has(key)) invoiceItems.set(key, []);
          const bucket = invoiceItems.get(key);
          if (!bucket.some(existing => existing.materialKey === item.materialKey && norm(existing.sku) === norm(item.sku))) {
            bucket.push(item);
          }
        }
      }
    } catch (error) {
      invoiceErrors.push(file.name + ': ' + (error.message || 'gagal dibaca.'));
    }
  }

  if (!invoiceSheetsRead || !invoiceItems.size) {
    throw new Error(invoiceErrors.join('\\n') || 'Tidak ditemukan sheet Invoice Adonia dengan kolom NO, SKU, MATERIAL CODE dan nomor invoice.');
  }

  const outputs = [];
  const details = [];
  const issues = [...invoiceErrors];
  let adoniaAjuCount = 0;
  let itemMatched = 0;
  let itemUnmatched = 0;
  let itemSkipped = 0;
  let outputRows = 0;

  for (const file of ceisaFiles) {
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const headerRows = rowsOf(workbook, 'HEADER', file.name);
      const entityRows = rowsOf(workbook, 'ENTITAS', file.name);
      const documentRows = rowsOf(workbook, 'DOKUMEN', file.name);
      const goodsRows = rowsOf(workbook, 'BARANG', file.name);

      const entitiesByAju = new Map();
      for (const row of entityRows.slice(1)) {
        const aju = clean(row[0]);
        if (!aju) continue;
        if (!entitiesByAju.has(aju)) entitiesByAju.set(aju, []);
        entitiesByAju.get(aju).push(row);
      }

      const documentsByAju = new Map();
      for (const row of documentRows.slice(1)) {
        const aju = clean(row[0]);
        if (!aju) continue;
        if (!documentsByAju.has(aju)) documentsByAju.set(aju, []);
        documentsByAju.get(aju).push(row);
      }

      const adoniaByAju = new Map();
      for (const header of headerRows.slice(1)) {
        const aju = clean(header[0]);
        if (!aju || clean(header[1]) !== '27') continue;

        const tanggalDaftar = formatCeisaDate(header[94]);
        const dateMatch = tanggalDaftar.match(/^(\\d{4})-(\\d{2})-/);
        const dateKey = dateMatch ? dateMatch[1] + '-' + dateMatch[2] : '';
        if (period?.year && period?.month && dateKey !== String(period.year) + '-' + String(period.month).padStart(2, '0')) continue;

        const entities = entitiesByAju.get(aju) || [];
        const chinli = entities.some(row => clean(row[2]) === '3' && norm(row[5]) === 'CHINLIPLASTICMATERIALSINDONESIA');
        const adonia = entities.some(row => clean(row[2]) === '8' && /ADONIA/i.test(clean(row[5])));
        if (!chinli || !adonia) continue;

        const documents = documentsByAju.get(aju) || [];
        const invoiceDocument = documents.find(row => clean(row[2]) === '380') || documents.find(row => clean(row[2]) === '217');
        const invoiceNo = clean(invoiceDocument?.[3]).toUpperCase();
        if (!invoiceNo) {
          issues.push(file.name + ' · ' + aju + ': nomor invoice tidak ditemukan pada DOKUMEN (kode 380/217). Baris barang tetap dimasukkan tanpa penambahan SKU.');
        }
        adoniaByAju.set(aju, {
          invoiceNo,
          nomorDaftar: clean(header[93]),
          tanggalDaftar
        });
        adoniaAjuCount++;
      }

      const rows = [];
      let changed = 0;
      for (let rowIndex = 1; rowIndex < goodsRows.length; rowIndex++) {
        const row = goodsRows[rowIndex] || [];
        const aju = clean(row[0]);
        const ajuInfo = adoniaByAju.get(aju);
        if (!ajuInfo) continue;

        const lineNo = clean(row[1]);
        const materialCode = clean(row[3]);
        const invoiceKey = ajuInfo.invoiceNo && /^\\d+$/.test(lineNo)
          ? ajuInfo.invoiceNo + '|' + String(Number(lineNo))
          : '';
        const candidates = invoiceKey ? (invoiceItems.get(invoiceKey) || []) : [];
        const valid = candidates.filter(item => item.materialKey === norm(materialCode));
        let sku = '';
        let status = 'INVOICE ITEM NOT FOUND';
        let note = 'Nomor invoice atau nomor urut item tidak ditemukan pada Invoice.';

        if (valid.length === 1) {
          sku = valid[0].sku;
          status = 'MATCHED';
          note = 'SKU Invoice cocok dengan Nomor Invoice, Nomor Urut Item, dan Material Code.';
          itemMatched++;
        } else if (valid.length > 1) {
          status = 'DUPLICATE INVOICE ITEM';
          note = 'Lebih dari satu baris Invoice cocok; SKU tidak ditambahkan.';
          itemUnmatched++;
        } else if (candidates.length) {
          status = 'MATERIAL CODE MISMATCH';
          note = 'Material Code Invoice berbeda dari Kode Barang CEISA; SKU tidak ditambahkan.';
          itemUnmatched++;
        } else {
          itemUnmatched++;
        }

        const existingSpecification = clean(row[8]);
        const finalSpecification = mergeSpecification(existingSpecification, sku);
        if (sku && finalSpecification !== existingSpecification) changed++;
        else if (sku) itemSkipped++;

        rows.push({
          'Nomor AJU': aju,
          'Nomor Daftar': ajuInfo.nomorDaftar,
          'Tanggal Daftar': ajuInfo.tanggalDaftar,
          'Kode Barang': clean(row[3]),
          'Uraian Barang': clean(row[4]),
          'Spesifikasi Lain': finalSpecification,
          'Jenis Satuan': clean(row[9]),
          'Kuantiti': numberValue(row[10])
        });

        details.push({
          file: file.name,
          aju,
          invoiceNo: ajuInfo.invoiceNo,
          lineNo,
          materialCode,
          sku,
          status,
          note
        });
      }

      rows.sort((a, b) =>
        a['Nomor AJU'].localeCompare(b['Nomor AJU'], 'en', { numeric: true }) ||
        String(a['Nomor Daftar']).localeCompare(String(b['Nomor Daftar']), 'en', { numeric: true }) ||
        a['Kode Barang'].localeCompare(b['Kode Barang'], 'en', { numeric: true })
      );

      const outputWorkbook = XLSX.utils.book_new();
      const outputSheet = XLSX.utils.json_to_sheet(rows, { header: ADONIA_OUTPUT_HEADERS });
      outputSheet['!cols'] = [24, 16, 16, 22, 48, 42, 16, 16].map(wch => ({ wch }));
      if (outputSheet['!ref']) outputSheet['!autofilter'] = { ref: outputSheet['!ref'] };
      XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'BARANG');

      outputs.push({
        name: file.name.replace(/\\.xlsx$/i, '') + '_ADONIA_BARANG.xlsx',
        data: XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'array' }),
        changed,
        rows: rows.length,
        rowsPreview: rows
      });
      outputRows += rows.length;
    } catch (error) {
      issues.push(file.name + ': ' + (error.message || 'gagal diproses.'));
    }
  }

  if (!outputs.length) throw new Error(issues.join('\\n') || 'Tidak ada file CEISA yang dapat diproses.');
  if (!outputRows) issues.push('Tidak ditemukan baris BARANG untuk transaksi Adonia pada periode yang dipilih.');

  return {
    outputs,
    details,
    issues,
    summary: {
      invoiceFiles: invoiceFiles.length,
      invoiceSheetsRead,
      adoniaAjuCount,
      itemMatched,
      itemUnmatched,
      itemSkipped,
      outputFiles: outputs.length,
      outputRows
    }
  };
}
