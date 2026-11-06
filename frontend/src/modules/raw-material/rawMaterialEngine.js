import * as XLSX from 'xlsx';

export const BAHANBAKU_HEADERS = [
  'NOMOR AJU','SERI BARANG','SERI BAHAN BAKU','KODE ASAL BAHAN BAKU','HS','KODE BARANG','URAIAN','MEREK','TIPE','UKURAN','SPESIFIKASI LAIN','KODE SATUAN','JUMLAH SATUAN','KODE KEMASAN','JUMLAH KEMASAN','KODE DOKUMEN ASAL','KODE KANTOR ASAL','NOMOR DAFTAR ASAL','TANGGAL DAFTAR ASAL','NOMOR AJU ASAL','SERI BARANG ASAL','NETTO','BRUTO','VOLUME','CIF','CIF RUPIAH','NDPBM','HARGA PENYERAHAN','HARGA PEROLEHAN','NILAI JASA','SERI IZIN','VALUTA','KODE BKC','KODE KOMODITI BKC','KODE SUB KOMODITI BKC','FLAG TIS','ISI PER KEMASAN','JUMLAH DILEKATKAN','JUMLAH PITA CUKAI','HJE CUKAI','TARIF CUKAI'
];
export const BAHANBAKUTARIF_HEADERS = [
  'NOMOR AJU','SERI BARANG','SERI BAHAN BAKU','KODE ASAL BAHAN BAKU','KODE PUNGUTAN','KODE TARIF','TARIF','KODE FASILITAS','TARIF FASILITAS','NILAI BAYAR','NILAI FASILITAS','NILAI SUDAH DILUNASI','KODE SATUAN','JUMLAH SATUAN','FLAG BMT SEMENTARA','KODE KOMODITI CUKAI','KODE SUB KOMODITI CUKAI','FLAG TIS','FLAG PELEKATAN','KODE KEMASAN','KODE SUB KOMODITI CUKAI','FLAG TIS','FLAG PELEKATAN','KODE KEMASAN','JUMLAH KEMASAN'
];
const col = (row, index) => row?.[index] ?? '';
const num = value => { if (value === '' || value == null) return 0; const n = Number(value); return Number.isFinite(n) ? n : 0; };
const clean = value => value == null ? '' : value;
const excelDate = value => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0,10);
  if (typeof value === 'number') { const d = XLSX.SSF.parse_date_code(value); if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`; }
  return clean(value);
};
function readSheet(workbook, name) {
  const sheet = workbook.Sheets[name] || workbook.Sheets[name.toUpperCase()];
  if (!sheet) throw new Error(`Sheet ${name} tidak ditemukan.`);
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false });
}
function firstDataRow(rows) { return rows.findIndex((r, i) => i > 0 && r.some(v => v !== '' && v != null)); }

export async function parseCeisaFile(file) {
  if (!/\.xlsx$/i.test(file.name)) throw new Error(`${file.name}: hanya file .xlsx yang didukung.`);
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const headerRows = readSheet(wb, 'HEADER');
  const itemRows = readSheet(wb, 'BARANG');
  const hi = firstDataRow(headerRows);
  if (hi < 0) throw new Error(`${file.name}: sheet HEADER tidak memiliki baris transaksi.`);
  const header = headerRows[hi];
  const aju = clean(col(header, 0));
  const docCode = String(col(header, 1)).trim();
  if (!aju) throw new Error(`${file.name}: NOMOR AJU pada HEADER kolom A kosong.`);
  if (!['27','23','40'].includes(docCode)) throw new Error(`${file.name}: kode dokumen ${docCode || '(kosong)'} belum didukung. Sistem mendukung BC 2.7, BC 2.3, dan BC 4.0.`);
  const ndpbm = docCode === '40' ? 1 : num(col(header, 74));
  const start = firstDataRow(itemRows);
  if (start < 0) throw new Error(`${file.name}: sheet BARANG tidak memiliki baris barang.`);
  const items = itemRows.slice(start).filter(r => r.some(v => v !== '' && v != null)).filter(r => !col(r,0) || String(col(r,0)).trim() === String(aju).trim());
  if (!items.length) throw new Error(`${file.name}: tidak ada barang yang cocok dengan NOMOR AJU ${aju}.`);
  return {
    fileName: file.name, aju, docCode, docType: docCode === '27' ? 'BC 2.7' : docCode === '23' ? 'BC 2.3' : 'BC 4.0',
    kodeKantor: clean(col(header,2)), nomorDaftar: clean(col(header,93)), tanggalDaftar: excelDate(col(header,94)), ndpbm,
    items: items.map((r, index) => {
      const originalQty = num(col(r,10));
      const sourceCif = num(col(r,25));
      const cifUnit = originalQty ? sourceCif / originalQty : 0;
      return {
        id: `${aju}-${clean(col(r,1)) || index + 1}-${index}`,
        sourceAju: clean(col(r,0)) || aju, sourceSeries: clean(col(r,1)) || index + 1, series: clean(col(r,1)) || index + 1,
        hs: clean(col(r,2)), kodeBarang: clean(col(r,3)), uraian: clean(col(r,4)), merek: clean(col(r,5)), tipe: clean(col(r,6)), ukuran: clean(col(r,7)), spesifikasi: clean(col(r,8)), kodeSatuan: clean(col(r,9)),
        originalQty, quantity: originalQty, cifUnit, sourceCif, sourceHargaPenyerahan: num(col(r,33)), ndpbm
      };
    })
  };
}

export function buildBahanBakuRows(doc, items) {
  return items.map(item => {
    const cif = item.cifUnit * num(item.quantity);
    const cifRupiah = cif * doc.ndpbm;
    const row = Object.fromEntries(BAHANBAKU_HEADERS.map(h => [h, '']));
    Object.assign(row, {
      'NOMOR AJU': doc.aju, 'SERI BARANG': item.series, 'SERI BAHAN BAKU': item.seriesBahan ?? 1,
      'KODE ASAL BAHAN BAKU': doc.docCode === '40' ? 1 : 0, 'HS': item.hs, 'KODE BARANG': item.kodeBarang,
      'URAIAN': item.uraian, 'MEREK': item.merek, 'TIPE': item.tipe, 'UKURAN': item.ukuran, 'SPESIFIKASI LAIN': item.spesifikasi,
      'KODE SATUAN': item.kodeSatuan, 'JUMLAH SATUAN': num(item.quantity), 'KODE DOKUMEN ASAL': doc.docCode,
      'KODE KANTOR ASAL': doc.kodeKantor, 'NOMOR DAFTAR ASAL': doc.nomorDaftar, 'TANGGAL DAFTAR ASAL': doc.tanggalDaftar,
      'NOMOR AJU ASAL': doc.aju, 'SERI BARANG ASAL': item.sourceSeries ?? item.series,
      'VOLUME': doc.docCode === '40'
        ? (num(item.originalQty) ? num(item.sourceHargaPenyerahan) / num(item.originalQty) : 0)
        : num(item.cifUnit),
      'CIF': cif, 'CIF RUPIAH': cifRupiah,
      'NDPBM': doc.ndpbm, 'HARGA PENYERAHAN': doc.docCode === '40' ? item.sourceHargaPenyerahan : cifRupiah
    });
    return row;
  });
}

export function buildTarifRows(doc, items, settings) {
  const tariffs = doc.docCode === '40'
    ? [{ code: 'PPNLOKAL', rate: settings.ppnLokal }]
    : [{ code: 'BM', rate: settings.bm }, { code: 'PPN', rate: settings.ppn }, { code: 'PPH', rate: settings.pph }];
  return items.flatMap(item => tariffs.map(t => {
    const row = Object.fromEntries(BAHANBAKUTARIF_HEADERS.map(h => [h, '']));
    Object.assign(row, {
      'NOMOR AJU': doc.aju, 'SERI BARANG': item.series, 'SERI BAHAN BAKU': item.seriesBahan ?? 1,
      'KODE ASAL BAHAN BAKU': doc.docCode === '40' ? '1' : '0', 'KODE PUNGUTAN': t.code,
      'KODE TARIF': settings.kodeTarif, 'TARIF': num(t.rate), 'KODE FASILITAS': settings.kodeFasilitas,
      'TARIF FASILITAS': num(settings.tarifFasilitas)
    });
    return row;
  }));
}

export function rowsToAoa(headers, rows) { return [headers, ...rows.map(row => headers.map(h => row[h] ?? ''))]; }
