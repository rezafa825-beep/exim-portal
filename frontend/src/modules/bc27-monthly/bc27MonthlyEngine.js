import * as XLSX from 'xlsx';

export const DATABASE_HEADERS = [
  'Nomor AJU', 'Nomor Daftar', 'Tanggal Daftar', 'Entitas Perusahaan Tujuan',
  'HS Code', 'Kode Barang', 'Spesifikasi Lain', 'Uraian Barang', 'Jumlah Satuan', 'Jenis Satuan', 'Kategori Barang'
];
export const OUTPUT_HEADERS = [
  'No', 'Identitas Barang', 'Uraian Barang', 'Jumlah Satuan', 'Jenis Satuan', 'Kategori Barang'
];

export const CATEGORY_MASTER = {
  '1': 'Hasil Produksi',
  '2': 'Bahan Baku',
  '3': 'Barang Modal',
  '4': 'Peralatan Kantor',
  '5': 'Sisa Dari Proses Produksi/Limbah (Waste/Scrap) dan/atau Sisa',
  '6': 'Barang Yang Ditimbun Untuk Dijual',
  '7': 'Barang Yang Dipamerkan Untuk Dijual',
  '8': 'Barang Lainnya'
};

// Identifikasi BC 2.7 KELUAR mengikuti aturan modul WhatsApp Report:
// kode entitas 3 harus merupakan Chinli. Nama tidak menghapus awalan "PT.";
// ketidaksesuaian master data akan ditandai sebagai bukan dokumen keluar.
const CHINLI_ENTITY_NAME = 'CHINLI PLASTIC MATERIALS INDONESIA';
const value = (row, index) => row?.[index] ?? '';
const text = v => v == null ? '' : String(v).trim();
const numberValue = v => {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const s = String(v).trim().replace(/\s/g, '');
  if (!s) return 0;
  const n = Number(s.replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
};

function excelDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = text(v);
  if (!s) return '';
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  return s;
}

function sheetRows(wb, name, fileName) {
  const sheet = wb.Sheets[name];
  if (!sheet) throw new Error(`${fileName}: sheet ${name} tidak ditemukan.`);
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false });
}

function isDataRow(row) {
  return row?.some(v => v !== '' && v != null);
}

function rowsAfterHeader(rows) {
  // CEISA menggunakan baris pertama sebagai nama kolom.
  return rows.slice(1).filter(isDataRow);
}

function isChinliOutgoing(aju, entitiesByAju) {
  const entities = entitiesByAju.get(aju) || [];
  const entity3 = entities.find(r => text(value(r, 2)) === '3');
  return text(value(entity3, 5)).toUpperCase() === CHINLI_ENTITY_NAME;
}

function getDestination(aju, entitiesByAju) {
  const entities = entitiesByAju.get(aju) || [];
  const entity8 = entities.find(r => text(value(r, 2)) === '8');
  return text(value(entity8, 5));
}

export async function parseMonthlyBc27Files(files, period) {
  if (!files?.length) throw new Error('Pilih minimal satu file Excel CEISA.');
  const seenAju = new Set();
  const database = [];
  const errors = [];
  const notices = [];
  let totalHeaderDocs = 0;
  let nonOutgoingDocs = 0;
  let duplicateDocs = 0;
  let outsidePeriodDocs = 0;
  let outgoingDocs = 0;

  for (const file of files) {
    if (!/\.xlsx$/i.test(file.name)) {
      errors.push(`${file.name}: format tidak didukung. Gunakan .xlsx.`);
      continue;
    }
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const headers = rowsAfterHeader(sheetRows(wb, 'HEADER', file.name));
      const entities = rowsAfterHeader(sheetRows(wb, 'ENTITAS', file.name));
      const goods = rowsAfterHeader(sheetRows(wb, 'BARANG', file.name));
      const entitiesByAju = new Map();
      for (const row of entities) {
        const aju = text(value(row, 0));
        if (!aju) continue;
        if (!entitiesByAju.has(aju)) entitiesByAju.set(aju, []);
        entitiesByAju.get(aju).push(row);
      }
      const goodsByAju = new Map();
      for (const row of goods) {
        const aju = text(value(row, 0));
        if (!aju) continue;
        if (!goodsByAju.has(aju)) goodsByAju.set(aju, []);
        goodsByAju.get(aju).push(row);
      }

      for (const header of headers) {
        const aju = text(value(header, 0));
        const docCode = text(value(header, 1));
        if (!aju || docCode !== '27') continue;
        totalHeaderDocs += 1;
        if (seenAju.has(aju)) {
          duplicateDocs += 1;
          continue;
        }
        seenAju.add(aju);
        if (!isChinliOutgoing(aju, entitiesByAju)) {
          nonOutgoingDocs += 1;
          continue;
        }
        const nomorDaftar = text(value(header, 93));
        const tanggalDaftar = excelDate(value(header, 94));
        const [year, month] = (tanggalDaftar || '').split('-');
        if (period?.year && period?.month && (year !== String(period.year) || month !== String(period.month).padStart(2, '0'))) {
          outsidePeriodDocs += 1;
          continue;
        }
        outgoingDocs += 1;
        const tujuan = getDestination(aju, entitiesByAju);
        const itemRows = goodsByAju.get(aju) || [];
        if (!itemRows.length) {
          notices.push(`${aju}: tidak ditemukan baris barang pada sheet BARANG.`);
          continue;
        }
        for (let i = 0; i < itemRows.length; i++) {
          const item = itemRows[i];
          const categoryCode = text(value(item, 48));
          const categoryName = CATEGORY_MASTER[categoryCode] || (categoryCode ? `Kode kategori tidak dikenal (${categoryCode})` : '');
          const kodeBarang = text(value(item, 3));
          const spesifikasiLain = text(value(item, 8));
          const uraian = text(value(item, 4));
          const unit = text(value(item, 9));
          const qty = numberValue(value(item, 10));
          if (!kodeBarang) notices.push(`${aju} · Baris barang ${i + 2}: Kode Barang kosong.`);
          if (tujuan.toUpperCase().includes('ADONIA') && !spesifikasiLain) notices.push(`${aju} · Baris barang ${i + 2}: Spesifikasi Lain Adonia kosong.`);
          if (!unit) notices.push(`${aju} · Baris barang ${i + 2}: Jenis Satuan kosong.`);
          if (!categoryCode || !CATEGORY_MASTER[categoryCode]) notices.push(`${aju} · Baris barang ${i + 2}: Kategori barang ${categoryCode || 'kosong'} perlu diperiksa.`);
          database.push({
            'Nomor AJU': aju,
            'Nomor Daftar': nomorDaftar,
            'Tanggal Daftar': tanggalDaftar,
            'Entitas Perusahaan Tujuan': tujuan || 'Tidak ditemukan (kode entitas 8)',
            'HS Code': text(value(item, 2)),
            'Kode Barang': kodeBarang,
            'Spesifikasi Lain': spesifikasiLain,
            'Uraian Barang': uraian,
            'Jumlah Satuan': qty,
            'Jenis Satuan': unit,
            'Kategori Barang': categoryName
          });
        }
      }
    } catch (error) {
      errors.push(error?.message ? `${file.name}: ${error.message}` : `${file.name}: gagal dibaca.`);
    }
  }

  if (!database.length && errors.length) throw new Error(errors.join('\n'));
  const output = buildMonthlyOutput(database);
  return {
    database,
    output,
    summary: { files: files.length, totalHeaderDocs, outgoingDocs, nonOutgoingDocs, duplicateDocs, outsidePeriodDocs, databaseRows: database.length, outputRows: output.length },
    errors,
    notices
  };
}

export function buildMonthlyOutput(database) {
  const grouped = new Map();
  for (const row of database) {
    const isAdonia = text(row['Entitas Perusahaan Tujuan']).toUpperCase().includes('ADONIA');
    // Adonia memakai Spesifikasi Lain sebagai identitas barang; perusahaan lain memakai Kode Barang.
    const identity = isAdonia ? text(row['Spesifikasi Lain']) : text(row['Kode Barang']);
    const unit = text(row['Jenis Satuan']);
    const category = text(row['Kategori Barang']);
    // Untuk Adonia, spesifikasi wajib menjadi bagian kunci. Baris tanpa identitas tidak digabung.
    const key = identity
      ? `${isAdonia ? 'ADONIA' : 'KODE'}\u001f${identity}\u001f${unit}\u001f${category}`
      : `NO-IDENTITY\u001f${row['Nomor AJU']}\u001f${grouped.size}`;
    if (!grouped.has(key)) grouped.set(key, {
      identity,
      description: text(row['Uraian Barang']),
      quantity: 0,
      unit,
      category,
      descriptions: new Set(),
      isAdonia
    });
    const target = grouped.get(key);
    target.quantity += numberValue(row['Jumlah Satuan']);
    if (text(row['Uraian Barang'])) target.descriptions.add(text(row['Uraian Barang']));
  }
  return [...grouped.values()]
    .sort((a,b) => a.identity.localeCompare(b.identity, 'id', { numeric: true }) || a.unit.localeCompare(b.unit) || a.category.localeCompare(b.category))
    .map((row, index) => ({
      'No': index + 1,
      'Identitas Barang': row.identity,
      'Uraian Barang': row.descriptions.size > 1 ? `${row.description} [PERIKSA: uraian berbeda]` : row.description,
      'Jumlah Satuan': row.quantity,
      'Jenis Satuan': row.unit,
      'Kategori Barang': row.category
    }));
}

export function exportMonthlyWorkbook(database, output, period) {
  const wb = XLSX.utils.book_new();
  const dbSheet = XLSX.utils.json_to_sheet(database, { header: DATABASE_HEADERS });
  const outSheet = XLSX.utils.json_to_sheet(output, { header: OUTPUT_HEADERS });
  dbSheet['!cols'] = [18,14,14,34,14,20,28,48,16,14,42].map(wch => ({ wch }));
  outSheet['!cols'] = [8,32,55,18,16,42].map(wch => ({ wch }));
  if (dbSheet['!ref']) dbSheet['!autofilter'] = { ref: dbSheet['!ref'] };
  if (outSheet['!ref']) outSheet['!autofilter'] = { ref: outSheet['!ref'] };
  XLSX.utils.book_append_sheet(wb, dbSheet, 'DATABASE');
  XLSX.utils.book_append_sheet(wb, outSheet, 'OUTPUT');
  const month = String(period.month).padStart(2, '0');
  XLSX.writeFile(wb, `BC27_MONTHLY_${period.year}_${month}.xlsx`);
}
