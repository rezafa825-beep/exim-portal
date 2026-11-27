import * as XLSX from 'xlsx';

const REQUIRED_SHEETS = ['HEADER', 'ENTITAS', 'BARANG'];
const DOCUMENT_SHEETS = new Set(['ENTITAS', 'DOKUMEN', 'PENGANGKUT', 'KEMASAN', 'KONTAINER', 'KOMPONENBIAYA', 'PUNGUTAN', 'JAMINAN', 'BANKDEVISA', 'RESPON', 'VERSI']);
const ITEM_SHEETS = new Set(['BARANGTARIF', 'BARANGDOKUMEN', 'BARANGENTITAS', 'BARANGSPEKKHUSUS', 'BARANGVD']);
const OPTIONAL_SHEETS = [
  'DOKUMEN', 'PENGANGKUT', 'KEMASAN', 'KONTAINER', 'KOMPONENBIAYA',
  'BARANGTARIF', 'BARANGDOKUMEN', 'BARANGENTITAS', 'BARANGSPEKKHUSUS',
  'BARANGVD', 'BAHANBAKU', 'BAHANBAKUTARIF', 'BAHANBAKUDOKUMEN',
  'PUNGUTAN', 'JAMINAN', 'BANKDEVISA', 'VERSI', 'RESPON'
];

function normalizeText(value) {
  return String(value ?? '').trim();
}

function normalizeCompanyName(value) {
  return normalizeText(value).toUpperCase().replace(/\bPT\.?\s*/g, '').replace(/[^A-Z0-9]/g, '');
}

function safeValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString();
  return value;
}

function worksheetBounds(sheet) {
  let minRow = Infinity, minCol = Infinity, maxRow = -1, maxCol = -1;
  const addresses = Object.keys(sheet).filter(key => /^[A-Z]{1,3}[0-9]+$/.test(key));
  for (const address of addresses) {
    const cell = XLSX.utils.decode_cell(address);
    minRow = Math.min(minRow, cell.r); minCol = Math.min(minCol, cell.c);
    maxRow = Math.max(maxRow, cell.r); maxCol = Math.max(maxCol, cell.c);
  }
  if (sheet['!ref']) {
    try {
      const range = XLSX.utils.decode_range(sheet['!ref']);
      minRow = Math.min(minRow, range.s.r); minCol = Math.min(minCol, range.s.c);
      maxRow = Math.max(maxRow, range.e.r); maxCol = Math.max(maxCol, range.e.c);
    } catch { /* Invalid dimension metadata is ignored; cell addresses remain authoritative. */ }
  }
  if (maxRow < 0 || maxCol < 0) return null;
  return { minRow: Math.max(0, minRow), minCol: Math.max(0, minCol), maxRow, maxCol };
}

function extractWorksheet(workbook, sheetName, sourceId) {
  const actualName = workbook.SheetNames.find(name => name.toUpperCase() === sheetName);
  if (!actualName) return { name: sheetName, exists: false, headers: [], records: [] };
  const sheet = workbook.Sheets[actualName];
  const bounds = worksheetBounds(sheet);
  if (!bounds) return { name: actualName, exists: true, headers: [], records: [] };

  const matrix = [];
  for (let row = bounds.minRow; row <= bounds.maxRow; row++) {
    const values = [];
    for (let col = bounds.minCol; col <= bounds.maxCol; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      values.push(safeValue(sheet[address]?.v ?? null));
    }
    matrix.push(values);
  }

  const headerOffset = matrix.findIndex(row => row.filter(value => value !== null && normalizeText(value) !== '').length >= 2);
  if (headerOffset < 0) return { name: actualName, exists: true, headers: [], records: [] };
  const headers = matrix[headerOffset].map(value => normalizeText(value).toUpperCase());
  const records = [];

  for (let index = headerOffset + 1; index < matrix.length; index++) {
    const values = matrix[index];
    if (!values.some(value => value !== null && value !== '')) continue;
    const fields = headers.map((header, columnIndex) => ({
      name: header || `UNNAMED_${columnIndex + 1}`,
      column_index: bounds.minCol + columnIndex,
      column_letter: XLSX.utils.encode_col(bounds.minCol + columnIndex),
      value: values[columnIndex] ?? null
    }));
    const sourceRow = bounds.minRow + index + 1;
    const lookup = (name, occurrence = 0) => {
      const matches = fields.filter(field => field.name === name);
      return matches[occurrence]?.value ?? null;
    };
    const location = (name, occurrence = 0) => {
      const field = fields.filter(item => item.name === name)[occurrence];
      return field ? { source_id: sourceId, sheet: actualName, row: sourceRow, column: field.column_letter } : null;
    };
    records.push({
      source_location: { source_id: sourceId, sheet: actualName, row: sourceRow },
      raw_fields: fields,
      values: fields.map(field => field.value),
      get: lookup,
      locate: location
    });
  }
  return { name: actualName, exists: true, headers, records };
}

function groupByAju(records) {
  const grouped = new Map();
  const orphaned = [];
  for (const record of records || []) {
    const aju = normalizeText(record.get('NOMOR AJU'));
    if (!aju) { orphaned.push(record); continue; }
    if (!grouped.has(aju)) grouped.set(aju, []);
    grouped.get(aju).push(record);
  }
  return { grouped, orphaned };
}

function classifyDirection(entities, companyName) {
  const company = normalizeCompanyName(companyName);
  if (!company) return { direction: 'UNKNOWN', status: 'NEEDS_REVIEW', reason: 'Nama perusahaan pada proyek belum diisi.' };
  const matches = entities.filter(row => normalizeCompanyName(row.get('NAMA ENTITAS')) === company);
  const hasCode3 = matches.some(row => normalizeText(row.get('KODE ENTITAS')) === '3');
  const hasCode8 = matches.some(row => normalizeText(row.get('KODE ENTITAS')) === '8');
  if (hasCode3 && !hasCode8) return { direction: 'OUT', status: 'CLASSIFIED', reason: 'Perusahaan ditemukan pada entitas kode 3.' };
  if (hasCode8 && !hasCode3) return { direction: 'IN', status: 'CLASSIFIED', reason: 'Perusahaan ditemukan pada entitas kode 8.' };
  if (hasCode3 && hasCode8) return { direction: 'UNKNOWN', status: 'NEEDS_REVIEW', reason: 'Identitas perusahaan muncul pada entitas kode 3 dan 8.' };
  return { direction: 'UNKNOWN', status: 'NEEDS_REVIEW', reason: 'Identitas perusahaan tidak ditemukan pada entitas kode 3 atau 8.' };
}

function rawRows(rows) {
  return rows.map(row => ({
    source_location: row.source_location,
    raw_fields: row.raw_fields
  }));
}

function normalizedEntity(row) {
  return {
    series: row.get('SERI'),
    entity_code: row.get('KODE ENTITAS'),
    identity_type: row.get('KODE JENIS IDENTITAS'),
    identity_number: row.get('NOMOR IDENTITAS'),
    name: row.get('NAMA ENTITAS'),
    address: row.get('ALAMAT ENTITAS'),
    nib: row.get('NIB ENTITAS'),
    source_location: row.source_location,
    raw_fields: row.raw_fields
  };
}

function normalizedDocument(header, related, classification, sourceId) {
  const aju = normalizeText(header.get('NOMOR AJU'));
  const registrationNumber = normalizeText(header.get('NOMOR DAFTAR')) || null;
  const registrationDate = header.get('TANGGAL DAFTAR') || null;
  const entities = (related.ENTITAS || []).map(normalizedEntity);
  const currentData = {
    nomor_aju: aju,
    kode_dokumen: header.get('KODE DOKUMEN'),
    kode_kantor: header.get('KODE KANTOR'),
    kode_kantor_tujuan: header.get('KODE KANTOR TUJUAN'),
    nomor_daftar: registrationNumber,
    tanggal_daftar: registrationDate,
    tanggal_pernyataan: header.get('TANGGAL PERNYATAAN'),
    kode_valuta: header.get('KODE VALUTA'),
    ndpbm: header.get('NDPBM'),
    cif: header.get('CIF'),
    harga_penyerahan: header.get('HARGA_PENYERAHAN') ?? header.get('HARGA PENYERAHAN'),
    bruto: header.get('BRUTO'),
    netto: header.get('NETTO'),
    volume: header.get('VOLUME'),
    entities
  };
  const extension = {};
  for (const [sheetName, rows] of Object.entries(related)) {
    if (sheetName !== 'BARANG' && sheetName !== 'BAHANBAKU' && sheetName !== 'BAHANBAKUTARIF') {
      extension[sheetName.toLowerCase()] = rawRows(rows);
    }
  }
  return {
    document_type: classification.status === 'CLASSIFIED'
      ? `BC27_${classification.direction}`
      : 'BC27_UNKNOWN',
    aju_number: aju,
    document_number: null,
    registration_number: registrationNumber,
    dates: { registration: registrationDate, declaration: header.get('TANGGAL PERNYATAAN') || null },
    parties: { entities },
    source_refs: [sourceId],
    original_data: {
      header: { source_location: header.source_location, raw_fields: header.raw_fields },
      related: extension
    },
    current_data: currentData,
    extension: {
      classification,
      header_raw_fields: header.raw_fields,
      related: extension
    },
    source_locations: {
      aju_number: header.locate('NOMOR AJU'),
      registration_number: header.locate('NOMOR DAFTAR'),
      registration_date: header.locate('TANGGAL DAFTAR')
    }
  };
}

function buildItems(aju, barangRows, bahanRows, tariffRows, otherItemRows, rawMaterialDocumentRows, sourceId, documentId) {
  const items = [];
  const rawMaterials = [];
  const rawMaterialTariffs = [];
  for (const row of barangRows) {
    const sequence = normalizeText(row.get('SERI BARANG'));
    const relatedMaterials = bahanRows.filter(material =>
      normalizeText(material.get('SERI BARANG')) === sequence
    );
    const relatedItemDetails = otherItemRows.filter(detail =>
      normalizeText(detail.get('SERI BARANG')) === sequence
    );
    const itemId = `ITEM-${crypto.randomUUID()}`;
    const current = {
      seri_barang: row.get('SERI BARANG'),
      hs: row.get('HS'),
      item_code: row.get('KODE BARANG'),
      description: row.get('URAIAN'),
      brand: row.get('MEREK'),
      type: row.get('TIPE'),
      size: row.get('UKURAN'),
      specification: row.get('SPESIFIKASI LAIN'),
      unit: row.get('KODE SATUAN'),
      quantity: row.get('JUMLAH SATUAN'),
      package_code: row.get('KODE KEMASAN'),
      package_quantity: row.get('JUMLAH KEMASAN'),
      net_weight: row.get('NETTO'),
      gross_weight: row.get('BRUTO'),
      volume: row.get('VOLUME'),
      cif: row.get('CIF'),
      cif_rupiah: row.get('CIF RUPIAH'),
      ndpbm: row.get('NDPBM'),
      harga_penyerahan: row.get('HARGA PENYERAHAN'),
      kode_asal_barang: row.get('KODE ASAL BARANG')
    };
    items.push({
      item_id: itemId,
      document_id: documentId,
      item_sequence: row.get('SERI BARANG'),
      item_code: row.get('KODE BARANG'),
      description: row.get('URAIAN'),
      specification: row.get('SPESIFIKASI LAIN'),
      quantity: row.get('JUMLAH SATUAN'),
      unit: row.get('KODE SATUAN'),
      net_weight: row.get('NETTO'),
      gross_weight: row.get('BRUTO'),
      currency: null,
      item_value: row.get('CIF') ?? row.get('HARGA PENYERAHAN'),
      source_refs: [sourceId],
      source_locations: row.raw_fields.map(field => ({
        field: field.name, column: field.column_letter,
        source_id: sourceId, sheet: row.source_location.sheet, row: row.source_location.row
      })),
      original_data: current,
      current_data: current,
      extension: {
        raw_fields: row.raw_fields,
        related_item_details: rawRows(relatedItemDetails)
      },
      bahan_baku_count: relatedMaterials.length
    });

    for (const material of relatedMaterials) {
      const materialId = `BB-${crypto.randomUUID()}`;
      const materialSequence = normalizeText(material.get('SERI BAHAN BAKU'));
      const linkedDocuments = rawMaterialDocumentRows.filter(detail =>
        normalizeText(detail.get('SERI BARANG')) === sequence &&
        normalizeText(detail.get('SERI BAHAN BAKU')) === materialSequence
      );
      const linkedTariffs = tariffRows.filter(tariff =>
        normalizeText(tariff.get('SERI BARANG')) === sequence &&
        normalizeText(tariff.get('SERI BAHAN BAKU')) === materialSequence
      );
      const materialCurrent = {
        seri_bahan_baku: material.get('SERI BAHAN BAKU'),
        kode_asal_bahan_baku: material.get('KODE ASAL BAHAN BAKU'),
        hs: material.get('HS'),
        item_code: material.get('KODE BARANG'),
        description: material.get('URAIAN'),
        unit: material.get('KODE SATUAN'),
        quantity: material.get('JUMLAH SATUAN'),
        cif: material.get('CIF'),
        cif_rupiah: material.get('CIF RUPIAH'),
        ndpbm: material.get('NDPBM'),
        harga_penyerahan: material.get('HARGA PENYERAHAN'),
        harga_perolehan: material.get('HARGA PEROLEHAN')
      };
      rawMaterials.push({
        raw_material_id: materialId,
        document_id: documentId,
        item_id: itemId,
        item_sequence: row.get('SERI BARANG'),
        raw_material_sequence: material.get('SERI BAHAN BAKU'),
        origin_code: material.get('KODE ASAL BAHAN BAKU'),
        item_code: material.get('KODE BARANG'),
        description: material.get('URAIAN'),
        quantity: material.get('JUMLAH SATUAN'),
        unit: material.get('KODE SATUAN'),
        source_refs: [sourceId],
        source_locations: material.raw_fields.map(field => ({
          field: field.name, column: field.column_letter,
          source_id: sourceId, sheet: material.source_location.sheet, row: material.source_location.row
        })),
        original_data: materialCurrent,
        current_data: materialCurrent,
        extension: { raw_fields: material.raw_fields, related_documents: rawRows(linkedDocuments) }
      });

      for (const tariff of linkedTariffs) {
        const tariffCurrent = {
          kode_pungutan: tariff.get('KODE PUNGUTAN'),
          kode_tarif: tariff.get('KODE TARIF'),
          tarif: tariff.get('TARIF'),
          kode_fasilitas: tariff.get('KODE FASILITAS'),
          tarif_fasilitas: tariff.get('TARIF FASILITAS'),
          nilai_bayar: tariff.get('NILAI BAYAR'),
          nilai_fasilitas: tariff.get('NILAI FASILITAS'),
          nilai_sudah_dilunasi: tariff.get('NILAI SUDAH DILUNASI'),
          kode_satuan: tariff.get('KODE SATUAN'),
          jumlah_satuan: tariff.get('JUMLAH SATUAN')
        };
        rawMaterialTariffs.push({
          raw_material_tariff_id: `BBT-${crypto.randomUUID()}`,
          document_id: documentId,
          item_id: itemId,
          raw_material_id: materialId,
          item_sequence: row.get('SERI BARANG'),
          raw_material_sequence: material.get('SERI BAHAN BAKU'),
          tax_code: tariff.get('KODE PUNGUTAN'),
          tariff_code: tariff.get('KODE TARIF'),
          tariff: tariff.get('TARIF'),
          source_refs: [sourceId],
          source_location: tariff.source_location,
          original_data: tariffCurrent,
          current_data: tariffCurrent,
          extension: { raw_fields: tariff.raw_fields }
        });
      }
    }
  }
  return { items, rawMaterials, rawMaterialTariffs };
}

export async function parseBC27File(file, companyName) {
  if (!file || !/\.(xlsx|xls|xlsm)$/i.test(file.name)) throw new Error('Format file harus Excel .xlsx, .xls, atau .xlsm.');
  const sourceId = `SRC-${crypto.randomUUID()}`;
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, raw: true, dense: false });
  const sheetData = {};
  for (const name of ['HEADER', ...REQUIRED_SHEETS.slice(1), ...OPTIONAL_SHEETS]) {
    sheetData[name] = extractWorksheet(workbook, name, sourceId);
  }

  const missing = REQUIRED_SHEETS.filter(name => !sheetData[name]?.exists);
  if (missing.length) throw new Error(`Workbook tidak dikenali sebagai ekspor CEISA BC 2.7. Worksheet wajib tidak ditemukan: ${missing.join(', ')}.`);
  const headerRows = sheetData.HEADER.records;
  if (!headerRows.length || !sheetData.HEADER.headers.includes('NOMOR AJU')) {
    throw new Error('Worksheet HEADER tidak memiliki record atau kolom NOMOR AJU.');
  }
  if (!headerRows.every(row => normalizeText(row.get('KODE DOKUMEN')) === '27')) {
    throw new Error('Workbook berisi kode dokumen selain BC 2.7 atau kode dokumen tidak terbaca.');
  }

  const allRows = {};
  for (const [name, data] of Object.entries(sheetData)) allRows[name] = data.records || [];
  const grouped = {};
  const orphanRows = [];
  for (const [name, rows] of Object.entries(allRows)) {
    const result = groupByAju(rows);
    grouped[name] = result.grouped;
    if (result.orphaned.length) orphanRows.push({ sheet: name, count: result.orphaned.length, rows: result.orphaned.map(row => row.source_location.row) });
  }

  const seenAju = new Set();
  const documents = [];
  const issues = [];
  const headerAjus = [];
  for (const header of headerRows) {
    const aju = normalizeText(header.get('NOMOR AJU'));
    if (!aju) {
      issues.push({ severity: 'ERROR', code: 'MISSING_AJU', message: 'Record HEADER tidak memiliki Nomor Aju.', sheet: 'HEADER', row: header.source_location.row });
      continue;
    }
    if (seenAju.has(aju)) {
      issues.push({ severity: 'ERROR', code: 'DUPLICATE_AJU_IN_FILE', message: `Nomor Aju ${aju} muncul lebih dari sekali pada HEADER.`, sheet: 'HEADER', row: header.source_location.row });
      continue;
    }
    seenAju.add(aju); headerAjus.push(aju);

    const related = {};
    for (const name of DOCUMENT_SHEETS) related[name] = grouped[name]?.get(aju) || [];
    const barangRows = grouped.BARANG?.get(aju) || [];
    const bahanRows = grouped.BAHANBAKU?.get(aju) || [];
    const tariffRows = grouped.BAHANBAKUTARIF?.get(aju) || [];
    const rawMaterialDocumentRows = grouped.BAHANBAKUDOKUMEN?.get(aju) || [];
    const otherItemRows = [...ITEM_SHEETS].flatMap(name => grouped[name]?.get(aju) || []);
    const entities = related.ENTITAS || [];
    const classification = classifyDirection(entities, companyName);
    if (classification.status !== 'CLASSIFIED') {
      issues.push({ severity: 'WARNING', code: 'DIRECTION_REVIEW', message: classification.reason, document_number: aju });
    }
    if (!barangRows.length) {
      issues.push({ severity: 'ERROR', code: 'NO_ITEMS', message: `Dokumen ${aju} tidak memiliki record BARANG.`, document_number: aju });
    }
    const documentId = `DOC-${crypto.randomUUID()}`;
    const document = normalizedDocument(header, related, classification, sourceId);
    document.document_id = documentId;
    const built = buildItems(aju, barangRows, bahanRows, tariffRows, otherItemRows, rawMaterialDocumentRows, sourceId, documentId);
    documents.push({
      document,
      items: built.items,
      raw_materials: built.rawMaterials,
      raw_material_tariffs: built.rawMaterialTariffs,
      validation: {
        status: issues.some(issue => issue.severity === 'ERROR' && issue.document_number === aju) ? 'ERROR' : classification.status === 'CLASSIFIED' ? 'READY' : 'WARNING',
        item_count: built.items.length,
        raw_material_count: built.rawMaterials.length,
        raw_material_tariff_count: built.rawMaterialTariffs.length
      }
    });
  }

  for (const row of orphanRows) {
    issues.push({ severity: 'WARNING', code: 'ORPHAN_ROWS', message: `${row.count} baris pada worksheet ${row.sheet} tidak memiliki Nomor Aju dan tidak dihubungkan otomatis.`, rows: row.rows });
  }
  for (const sheetName of OPTIONAL_SHEETS) {
    if (!sheetData[sheetName]?.exists) {
      issues.push({ severity: 'INFO', code: 'OPTIONAL_SHEET_MISSING', message: `Worksheet opsional ${sheetName} tidak ditemukan.`, sheet: sheetName });
    } else if (!sheetData[sheetName].records.length) {
      issues.push({ severity: 'INFO', code: 'OPTIONAL_SHEET_EMPTY', message: `Worksheet opsional ${sheetName} kosong.`, sheet: sheetName });
    }
  }

  const digest = await globalThis.crypto.subtle.digest('SHA-256', buffer);
  const fileHash = Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2, '0')).join('');
  const safeFileName = file.name.replace(/[^a-z0-9._-]+/gi, '_');
  const source = {
    source_id: sourceId,
    source_type: 'CEISA',
    original_file_name: file.name,
    internal_path: `sources/ceisa/${sourceId}_${safeFileName}`,
    extension: file.name.split('.').pop().toLowerCase(),
    mime_type: file.type || null,
    file_size: file.size,
    sha256: fileHash,
    imported_at: new Date().toISOString(),
    import_status: 'STAGED',
    document_ids: documents.map(row => row.document.document_id),
    file_format: headerRows.length === 1 ? 'SINGLE_DOCUMENT' : 'MULTI_DOCUMENT',
    document_count: documents.length
  };

  return {
    source,
    source_bytes: new Uint8Array(buffer),
    documents,
    issues,
    file_format: source.file_format,
    workbook_sheets: workbook.SheetNames,
    summary: {
      document_count: documents.length,
      item_count: documents.reduce((sum, row) => sum + row.items.length, 0),
      raw_material_count: documents.reduce((sum, row) => sum + row.raw_materials.length, 0),
      raw_material_tariff_count: documents.reduce((sum, row) => sum + row.raw_material_tariffs.length, 0),
      error_count: issues.filter(issue => issue.severity === 'ERROR').length,
      warning_count: issues.filter(issue => issue.severity === 'WARNING').length
    }
  };
}
